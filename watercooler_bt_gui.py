#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Standalone PyQt5 application with a system tray icon for managing a liquid-cooling system over BLE,
showing connection status, controlling RGB lighting, and manual or automatic cooling control.

Features:
- Manual mode: direct sliders for fan, pump, and RGB controls.
- Automatic mode: editable fan curve and pump curve based on CPU/GPU temperatures.
- Higher-contrast curve axes and labels for dark theme readability.
- Minimize to tray, close to exit.

Uses PyQt5 + qasync + pythonnet + bleak.
"""
import sys
import asyncio
import atexit
import logging
import traceback
from pathlib import Path
from enum import IntEnum
from collections import deque
from statistics import median

from PyQt5 import QtWidgets, QtCore, QtGui
import clr
from bleak import BleakScanner, BleakClient
import qasync
from qasync import asyncSlot

import os
import json
import platform
import ctypes
import subprocess
import time
import threading
import base64
import hashlib
import hmac
from datetime import datetime
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# === UI STRINGS ===
UI = {
    'mode_manual': "手动模式",
    'mode_curve': "自动模式",
    'searching': "正在搜索水冷设备...",
    'select_prompt': "请选择设备并点击连接",
    'no_device': "未找到BLE水冷设备，正在重试...",
    'connecting': "正在连接 {}...",
    'connected': "已连接至 {}",
    'label_fan': "风扇功率 (%):",
    'label_pump': "水泵电压 (V):",
        'label_rgb': "RGB灯效:",
    'label_curve': "自动曲线:",
    'btn_connect': "连接",
    'btn_disconnect': "断开连接",
    'btn_apply_manual': "应用",
    'btn_apply_rgb': "应用RGB",
    'btn_apply_all': "应用全部",
    'btn_apply_curve': "启用自动模式",
    'tray_exit': "退出",
    'tray_show': "显示",
    'label_update_speed': "更新速度:",
    'label_auto_connect': "自动连接",
    'label_auto_start': "开机自启"
}

MAX_SAFE_FAN_PERCENT = 90
MAX_SAFE_FAN_DUTY = 229  # 约等于 255 的 90%，参考上游项目 v1.1.0 的安全限制
DEFAULT_MANUAL_FAN_PERCENT = 30
DEFAULT_UPDATE_INTERVAL_SEC = 5.0
DEFAULT_FAN_CURVE_POINTS = [(40, 0), (45, 35), (55, 60), (60, MAX_SAFE_FAN_PERCENT)]
DEFAULT_PUMP_CURVE_POINTS = [(40, 0), (45, 7), (55, 8), (60, 11)]
PUMP_CURVE_LEVELS = [0, 7, 8, 11]
DEFAULT_RGB_TEMP_THRESHOLDS = (40, 60)
DEFAULT_RGB_TEMP_COLORS = {'low': (255, 255, 255), 'mid': (0, 255, 0), 'high': (255, 0, 0)}
DEFAULT_AUTO_DEBOUNCE_SAMPLES = 3
DEFAULT_AUTO_HYSTERESIS_C = 2.0
DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC = 3.0
DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC = 3.0
DEFAULT_EXPORT_API_HOST = '127.0.0.1'
DEFAULT_EXPORT_API_PORT = 21977
MIN_CURVE_POINTS = 2
MAX_CURVE_POINTS = 8

LEGACY_FAN_CURVE_DEFAULTS = [
    [(40, 0), (50, 50), (60, MAX_SAFE_FAN_PERCENT)],
    [(40, 0), (50, 50), (60, 90)],
]
LEGACY_PUMP_CURVE_DEFAULTS = [
    [(40, 0), (50, 8), (60, 11)],
]


class NoWheelSlider(QtWidgets.QSlider):
    def wheelEvent(self, event):
        event.ignore()


class NoWheelComboBox(QtWidgets.QComboBox):
    def wheelEvent(self, event):
        event.ignore()


class NoWheelSpinBox(QtWidgets.QSpinBox):
    def wheelEvent(self, event):
        event.ignore()


def _points_equal(a, b):
    try:
        return [tuple(map(int, point)) for point in a] == [tuple(map(int, point)) for point in b]
    except Exception:
        return False


def migrate_curve_defaults_if_needed(fan_points, pump_points):
    fan = [tuple(point) for point in fan_points]
    pump = [tuple(point) for point in pump_points]
    if any(_points_equal(fan, legacy) for legacy in LEGACY_FAN_CURVE_DEFAULTS):
        fan = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
    if any(_points_equal(pump, legacy) for legacy in LEGACY_PUMP_CURVE_DEFAULTS):
        pump = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
    return fan, pump

def clamp_fan_duty(duty: int) -> int:
    duty = int(duty)
    return max(0, min(duty, MAX_SAFE_FAN_DUTY))

def clamp_curve_percent(percent: int) -> int:
    percent = int(percent)
    return max(0, min(percent, MAX_SAFE_FAN_PERCENT))


def clamp_pump_curve_value(value: int) -> int:
    value = int(value)
    return min(PUMP_CURVE_LEVELS, key=lambda candidate: abs(candidate - value))


def fan_percent_to_duty(percent: int) -> int:
    return clamp_fan_duty(round(clamp_curve_percent(percent) / 100.0 * 255))


def _normalize_curve_points(points, value_normalizer, default_points):
    try:
        normalized = []
        for temp, value in points:
            temp = int(min(max(int(temp), 20), 100))
            normalized.append((temp, value_normalizer(value)))
        normalized.sort(key=lambda item: item[0])
        fixed = []
        for idx, (temp, value) in enumerate(normalized):
            if idx > 0 and temp <= fixed[-1][0]:
                temp = min(100, fixed[-1][0] + 1)
            fixed.append((temp, value))
        if len(fixed) >= 2:
            return fixed
    except Exception:
        pass
    return [tuple(point) for point in default_points]


def normalize_fan_curve_points(points):
    return _normalize_curve_points(points, clamp_curve_percent, DEFAULT_FAN_CURVE_POINTS)


def normalize_pump_curve_points(points):
    return _normalize_curve_points(points, clamp_pump_curve_value, DEFAULT_PUMP_CURVE_POINTS)


def pump_curve_value_to_text(value):
    try:
        value = int(value)
    except Exception:
        value = 0
    if value <= 0:
        return '关闭'
    return PUMP_VOLTAGE_NAMES.get(value, f'{value}V')


def normalize_update_interval(sec):
    try:
        sec = float(sec)
    except Exception:
        sec = DEFAULT_UPDATE_INTERVAL_SEC
    return max(0.2, min(sec, 60.0))


def pump_display_to_enum(value):
    mapping = {7: PumpVoltage.V7, 8: PumpVoltage.V8, 11: PumpVoltage.V11}
    return mapping.get(int(value), PumpVoltage.V7)


def pump_enum_to_display(voltage):
    mapping = {PumpVoltage.V7: 7, PumpVoltage.V8: 8, PumpVoltage.V11: 11}
    try:
        return mapping.get(PumpVoltage(voltage), 7)
    except Exception:
        return 7


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def setup_logging():
    log_file = get_app_dir() / 'watercooler.log'
    try:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(threadName)s %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout),
            ],
        )
        logging.info('==== WaterCooler Manager start ====')
    except Exception:
        pass


def install_global_exception_hooks(loop=None):
    def _log_exception(prefix, exc_type, exc_value, exc_tb):
        try:
            logging.critical('%s\n%s', prefix, ''.join(traceback.format_exception(exc_type, exc_value, exc_tb)))
        except Exception:
            pass

    def _sys_excepthook(exc_type, exc_value, exc_tb):
        _log_exception('Unhandled exception', exc_type, exc_value, exc_tb)
        try:
            sys.__excepthook__(exc_type, exc_value, exc_tb)
        except Exception:
            pass

    def _threading_excepthook(args):
        _log_exception(f'Unhandled thread exception: {args.thread.name}', args.exc_type, args.exc_value, args.exc_traceback)

    sys.excepthook = _sys_excepthook
    try:
        threading.excepthook = _threading_excepthook
    except Exception:
        pass

    if loop is not None:
        def _asyncio_exception_handler(loop, context):
            try:
                logging.error('Asyncio exception: %s', context.get('message', ''))
                exc = context.get('exception')
                if exc is not None:
                    logging.error(''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
                else:
                    logging.error('Asyncio context: %r', context)
            except Exception:
                pass

        try:
            loop.set_exception_handler(_asyncio_exception_handler)
        except Exception:
            pass


def is_windows_admin() -> bool:
    if platform.system() != 'Windows':
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _show_admin_error(message: str):
    if platform.system() == 'Windows':
        try:
            ctypes.windll.user32.MessageBoxW(None, message, '水冷管理器', 0x10)
            return
        except Exception:
            pass
    print(message, file=sys.stderr)


def ensure_admin_rights() -> bool:
    if platform.system() != 'Windows' or is_windows_admin():
        return True

    try:
        if getattr(sys, 'frozen', False):
            executable = sys.executable
            parameters = subprocess.list2cmdline(sys.argv[1:])
        else:
            executable = sys.executable
            parameters = subprocess.list2cmdline([str(Path(__file__).resolve()), *sys.argv[1:]])

        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            'runas',
            executable,
            parameters,
            str(get_app_dir()),
            1,
        )
        if result <= 32:
            _show_admin_error('程序必须使用管理员权限运行。\n\n请在弹出的 UAC 窗口中选择“是”，或右键以管理员身份运行。')
        return False
    except Exception as exc:
        _show_admin_error(f'申请管理员权限失败：{exc}')
        return False



class ExportApiState:
    def __init__(self):
        self._lock = threading.Lock()
        self._snapshot = {
            'ok': True,
            'app': 'watercooler-manager',
            'connected': False,
            'mode': 'manual',
            'fan': {'percent': 0, 'text': '0%', 'is_off': True},
            'pump': {'voltage': 0, 'text': '关闭', 'is_off': True},
            'temperature': {'cpu_c': None, 'gpu_c': None, 'control_c': None},
            'device_name': None,
            'timestamp': int(time.time()),
        }

    def update(self, **fields):
        with self._lock:
            self._snapshot.update(fields)
            self._snapshot['timestamp'] = int(time.time())

    def snapshot(self):
        with self._lock:
            return json.loads(json.dumps(self._snapshot, ensure_ascii=False))


class WatercoolerApiServer:
    def __init__(self, state: ExportApiState, host: str = DEFAULT_EXPORT_API_HOST, port: int = DEFAULT_EXPORT_API_PORT):
        self.state = state
        self.host = host
        self.port = int(port)
        self._server = None
        self._thread = None

    def start(self):
        if self._server is not None:
            return

        state = self.state

        class _Handler(BaseHTTPRequestHandler):
            def _send_json(self, payload, status=200):
                body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
                self.send_response(status)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                if self.path in ('/', '/api', '/api/status'):
                    self._send_json(state.snapshot())
                elif self.path == '/api/health':
                    self._send_json({'ok': True, 'timestamp': int(time.time())})
                else:
                    self._send_json({'ok': False, 'error': 'not_found'}, status=404)

            def log_message(self, format, *args):
                return

        try:
            self._server = ThreadingHTTPServer((self.host, self.port), _Handler)
        except OSError as exc:
            print(f'Export API start failed on {self.host}:{self.port}: {exc}')
            self._server = None
            return

        self._thread = threading.Thread(target=self._server.serve_forever, name='WatercoolerExportApi', daemon=True)
        self._thread.start()

    def stop(self):
        if self._server is None:
            return
        try:
            self._server.shutdown()
            self._server.server_close()
        except Exception:
            pass
        self._server = None
        self._thread = None


class Settings:
    REGISTRY_KEY = r"Software\WaterCooler"
    CONFIG_FILE = str(get_app_dir() / "watercooler.json")

    def __init__(self):
        self.current_voltage = PumpVoltage.V7
        self.current_fan_speed = fan_percent_to_duty(DEFAULT_MANUAL_FAN_PERCENT)
        self.pump_is_off = False
        self.fan_is_off = False
        self.rgb_state = RGBMode.STATIC
        self.rgb_is_off = False
        self.rgb_color = (255, 0, 0)
        self.auto_start = False
        self.auto_connect = False
        self.fan_curve_points = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
        self.pump_curve_points = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
        self.selected_mode_index = 0
        self.auto_mode_enabled = False
        self.update_interval_sec = DEFAULT_UPDATE_INTERVAL_SEC
        self.theme_mode = 'system'
        self.rgb_temp_enabled = True
        self.rgb_temp_mode = RGBMode.STATIC
        self.rgb_temp_threshold_low = DEFAULT_RGB_TEMP_THRESHOLDS[0]
        self.rgb_temp_threshold_high = DEFAULT_RGB_TEMP_THRESHOLDS[1]
        self.rgb_temp_color_low = DEFAULT_RGB_TEMP_COLORS['low']
        self.rgb_temp_color_mid = DEFAULT_RGB_TEMP_COLORS['mid']
        self.rgb_temp_color_high = DEFAULT_RGB_TEMP_COLORS['high']
        self.auto_hysteresis_c = DEFAULT_AUTO_HYSTERESIS_C
        self.auto_debounce_samples = DEFAULT_AUTO_DEBOUNCE_SAMPLES
        self.auto_fan_min_toggle_interval_sec = DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC
        self.auto_pump_min_toggle_interval_sec = DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC
        self.export_api_enabled = False
        self.export_api_port = DEFAULT_EXPORT_API_PORT
        self.last_device_address = None
        self.last_device_name = None
        self.dingtalk_webhook_enabled = False
        self.dingtalk_webhook_url = ''
        self.dingtalk_webhook_secret = ''
        self.load()
        self.normalize()
        self._sync_autostart_if_needed()

    def load(self):
        # 仅从运行目录配置文件加载。若文件不存在，则使用代码内默认值。
        # 不再从 Windows 注册表回退读取旧配置，避免删除 watercooler.json 后
        # 又把历史曲线（例如 3 点曲线）恢复回来。
        self._load_from_file()

    def save(self):
        self.normalize()
        self._save_to_file()

    def normalize(self):
        try:
            self.current_fan_speed = clamp_fan_duty(self.current_fan_speed)
        except Exception:
            self.current_fan_speed = 150

        try:
            self.selected_mode_index = int(self.selected_mode_index)
        except Exception:
            self.selected_mode_index = 0
        self.selected_mode_index = 1 if self.selected_mode_index == 1 else 0
        self.auto_mode_enabled = bool(self.auto_mode_enabled)
        self.fan_curve_points = normalize_fan_curve_points(self.fan_curve_points)
        self.pump_curve_points = normalize_pump_curve_points(self.pump_curve_points)
        self.fan_curve_points, self.pump_curve_points = migrate_curve_defaults_if_needed(self.fan_curve_points, self.pump_curve_points)
        self.update_interval_sec = normalize_update_interval(self.update_interval_sec)

        try:
            self.current_voltage = PumpVoltage(int(self.current_voltage))
        except Exception:
            self.current_voltage = PumpVoltage.V7
        if self.current_voltage == PumpVoltage.V12:
            self.current_voltage = PumpVoltage.V11

        if self.theme_mode not in ('dark', 'light', 'system'):
            self.theme_mode = 'system'

        if self.rgb_is_off:
            self.rgb_state = RGBMode.OFF

        try:
            self.rgb_temp_mode = RGBMode(int(self.rgb_temp_mode))
        except Exception:
            self.rgb_temp_mode = RGBMode.STATIC
        if self.rgb_temp_mode not in (RGBMode.STATIC, RGBMode.BREATH):
            self.rgb_temp_mode = RGBMode.STATIC

        try:
            low = int(self.rgb_temp_threshold_low)
        except Exception:
            low = DEFAULT_RGB_TEMP_THRESHOLDS[0]
        try:
            high = int(self.rgb_temp_threshold_high)
        except Exception:
            high = DEFAULT_RGB_TEMP_THRESHOLDS[1]
        low = max(20, min(low, 95))
        high = max(low + 1, min(high, 100))
        self.rgb_temp_threshold_low = low
        self.rgb_temp_threshold_high = high

        def _normalize_color(value, fallback):
            try:
                if isinstance(value, (list, tuple)) and len(value) == 3:
                    return tuple(max(0, min(255, int(v))) for v in value)
            except Exception:
                pass
            return fallback

        self.rgb_temp_color_low = _normalize_color(self.rgb_temp_color_low, DEFAULT_RGB_TEMP_COLORS['low'])
        self.rgb_temp_color_mid = _normalize_color(self.rgb_temp_color_mid, DEFAULT_RGB_TEMP_COLORS['mid'])
        self.rgb_temp_color_high = _normalize_color(self.rgb_temp_color_high, DEFAULT_RGB_TEMP_COLORS['high'])
        self.rgb_temp_enabled = bool(self.rgb_temp_enabled)

        try:
            self.auto_hysteresis_c = int(self.auto_hysteresis_c)
        except Exception:
            self.auto_hysteresis_c = DEFAULT_AUTO_HYSTERESIS_C
        self.auto_hysteresis_c = max(1, min(int(self.auto_hysteresis_c), 5))

        try:
            self.auto_debounce_samples = int(self.auto_debounce_samples)
        except Exception:
            self.auto_debounce_samples = DEFAULT_AUTO_DEBOUNCE_SAMPLES
        self.auto_debounce_samples = 5 if self.auto_debounce_samples >= 5 else 3

        try:
            self.auto_fan_min_toggle_interval_sec = float(self.auto_fan_min_toggle_interval_sec)
        except Exception:
            legacy_interval = getattr(self, 'auto_min_toggle_interval_sec', DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC)
            self.auto_fan_min_toggle_interval_sec = legacy_interval
        self.auto_fan_min_toggle_interval_sec = max(0.0, min(float(self.auto_fan_min_toggle_interval_sec), 30.0))

        try:
            self.auto_pump_min_toggle_interval_sec = float(self.auto_pump_min_toggle_interval_sec)
        except Exception:
            legacy_interval = getattr(self, 'auto_min_toggle_interval_sec', DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC)
            self.auto_pump_min_toggle_interval_sec = legacy_interval
        self.auto_pump_min_toggle_interval_sec = max(0.0, min(float(self.auto_pump_min_toggle_interval_sec), 30.0))

        self.last_device_address = str(getattr(self, 'last_device_address', '') or '').strip() or None
        self.last_device_name = str(getattr(self, 'last_device_name', '') or '').strip() or None

        self.export_api_enabled = bool(getattr(self, 'export_api_enabled', False))
        try:
            self.export_api_port = int(getattr(self, 'export_api_port', DEFAULT_EXPORT_API_PORT))
        except Exception:
            self.export_api_port = DEFAULT_EXPORT_API_PORT
        self.export_api_port = max(1024, min(int(self.export_api_port), 65535))

        self.dingtalk_webhook_enabled = bool(getattr(self, 'dingtalk_webhook_enabled', False))
        self.dingtalk_webhook_url = str(getattr(self, 'dingtalk_webhook_url', '') or '').strip()
        self.dingtalk_webhook_secret = str(getattr(self, 'dingtalk_webhook_secret', '') or '').strip()

    def _load_from_registry(self):
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REGISTRY_KEY)
            self.current_voltage = PumpVoltage(winreg.QueryValueEx(key, "current_voltage")[0])
            self.current_fan_speed = winreg.QueryValueEx(key, "current_fan_speed")[0]
            self.pump_is_off = bool(winreg.QueryValueEx(key, "pump_is_off")[0])
            self.fan_is_off = bool(winreg.QueryValueEx(key, "fan_is_off")[0])
            self.rgb_state = RGBMode(winreg.QueryValueEx(key, "rgb_state")[0])
            self.rgb_is_off = bool(winreg.QueryValueEx(key, "rgb_is_off")[0])
            self.rgb_color = tuple(winreg.QueryValueEx(key, "rgb_color")[0])
            self.auto_start = bool(winreg.QueryValueEx(key, "auto_start")[0])
            self.auto_connect = bool(winreg.QueryValueEx(key, "auto_connect")[0])
            try:
                self.selected_mode_index = winreg.QueryValueEx(key, "selected_mode_index")[0]
            except Exception:
                self.selected_mode_index = 0
            try:
                self.auto_mode_enabled = bool(winreg.QueryValueEx(key, "auto_mode_enabled")[0])
            except Exception:
                self.auto_mode_enabled = False
            try:
                self.fan_curve_points = json.loads(winreg.QueryValueEx(key, "fan_curve_points")[0])
            except Exception:
                self.fan_curve_points = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
            try:
                self.pump_curve_points = json.loads(winreg.QueryValueEx(key, "pump_curve_points")[0])
            except Exception:
                self.pump_curve_points = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
            try:
                self.update_interval_sec = winreg.QueryValueEx(key, "update_interval_sec")[0]
            except Exception:
                self.update_interval_sec = DEFAULT_UPDATE_INTERVAL_SEC
            try:
                self.theme_mode = winreg.QueryValueEx(key, "theme_mode")[0]
            except Exception:
                self.theme_mode = 'system'
            try:
                self.auto_hysteresis_c = winreg.QueryValueEx(key, "auto_hysteresis_c")[0]
            except Exception:
                self.auto_hysteresis_c = DEFAULT_AUTO_HYSTERESIS_C
            try:
                self.auto_debounce_samples = winreg.QueryValueEx(key, "auto_debounce_samples")[0]
            except Exception:
                self.auto_debounce_samples = DEFAULT_AUTO_DEBOUNCE_SAMPLES
            try:
                self.auto_fan_min_toggle_interval_sec = winreg.QueryValueEx(key, "auto_fan_min_toggle_interval_sec")[0]
            except Exception:
                try:
                    self.auto_fan_min_toggle_interval_sec = winreg.QueryValueEx(key, "auto_min_toggle_interval_sec")[0]
                except Exception:
                    self.auto_fan_min_toggle_interval_sec = DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC
            try:
                self.auto_pump_min_toggle_interval_sec = winreg.QueryValueEx(key, "auto_pump_min_toggle_interval_sec")[0]
            except Exception:
                try:
                    self.auto_pump_min_toggle_interval_sec = winreg.QueryValueEx(key, "auto_min_toggle_interval_sec")[0]
                except Exception:
                    self.auto_pump_min_toggle_interval_sec = DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC
            winreg.CloseKey(key)
            return True
        except Exception:
            return False

    def _load_from_file(self):
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self.current_voltage = PumpVoltage(config['current_voltage'])
                self.current_fan_speed = config['current_fan_speed']
                self.pump_is_off = config['pump_is_off']
                self.fan_is_off = config['fan_is_off']
                self.rgb_state = RGBMode(config['rgb_state'])
                self.rgb_is_off = config['rgb_is_off']
                self.rgb_color = tuple(config['rgb_color'])
                self.auto_start = config['auto_start']
                self.auto_connect = config['auto_connect']
                self.selected_mode_index = config.get('selected_mode_index', 0)
                self.auto_mode_enabled = config.get('auto_mode_enabled', False)
                self.fan_curve_points = config.get('fan_curve_points', DEFAULT_FAN_CURVE_POINTS)
                self.pump_curve_points = config.get('pump_curve_points', DEFAULT_PUMP_CURVE_POINTS)
                self.update_interval_sec = config.get('update_interval_sec', DEFAULT_UPDATE_INTERVAL_SEC)
                self.theme_mode = config.get('theme_mode', 'system')
                self.rgb_temp_enabled = config.get('rgb_temp_enabled', False)
                self.rgb_temp_mode = config.get('rgb_temp_mode', int(RGBMode.STATIC))
                self.rgb_temp_threshold_low = config.get('rgb_temp_threshold_low', DEFAULT_RGB_TEMP_THRESHOLDS[0])
                self.rgb_temp_threshold_high = config.get('rgb_temp_threshold_high', DEFAULT_RGB_TEMP_THRESHOLDS[1])
                self.rgb_temp_color_low = tuple(config.get('rgb_temp_color_low', DEFAULT_RGB_TEMP_COLORS['low']))
                self.rgb_temp_color_mid = tuple(config.get('rgb_temp_color_mid', DEFAULT_RGB_TEMP_COLORS['mid']))
                self.rgb_temp_color_high = tuple(config.get('rgb_temp_color_high', DEFAULT_RGB_TEMP_COLORS['high']))
                self.auto_hysteresis_c = config.get('auto_hysteresis_c', DEFAULT_AUTO_HYSTERESIS_C)
                self.auto_debounce_samples = config.get('auto_debounce_samples', DEFAULT_AUTO_DEBOUNCE_SAMPLES)
                legacy_min_toggle = config.get('auto_min_toggle_interval_sec', DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC)
                self.auto_fan_min_toggle_interval_sec = config.get('auto_fan_min_toggle_interval_sec', legacy_min_toggle)
                self.auto_pump_min_toggle_interval_sec = config.get('auto_pump_min_toggle_interval_sec', legacy_min_toggle)
                self.export_api_enabled = config.get('export_api_enabled', False)
                self.export_api_port = config.get('export_api_port', DEFAULT_EXPORT_API_PORT)
                self.last_device_address = config.get('last_device_address')
                self.last_device_name = config.get('last_device_name')
                self.dingtalk_webhook_enabled = config.get('dingtalk_webhook_enabled', False)
                self.dingtalk_webhook_url = config.get('dingtalk_webhook_url', '')
                self.dingtalk_webhook_secret = config.get('dingtalk_webhook_secret', '')
            return True
        except Exception:
            return False

    def _save_to_file(self):
        try:
            config = {
                'current_voltage': int(self.current_voltage),
                'current_fan_speed': self.current_fan_speed,
                'pump_is_off': self.pump_is_off,
                'fan_is_off': self.fan_is_off,
                'rgb_state': int(self.rgb_state),
                'rgb_is_off': self.rgb_is_off,
                'rgb_color': self.rgb_color,
                'auto_start': self.auto_start,
                'auto_connect': self.auto_connect,
                'selected_mode_index': self.selected_mode_index,
                'auto_mode_enabled': self.auto_mode_enabled,
                'fan_curve_points': self.fan_curve_points,
                'pump_curve_points': self.pump_curve_points,
                'update_interval_sec': self.update_interval_sec,
                'theme_mode': self.theme_mode,
                'rgb_temp_enabled': self.rgb_temp_enabled,
                'rgb_temp_mode': int(self.rgb_temp_mode),
                'rgb_temp_threshold_low': self.rgb_temp_threshold_low,
                'rgb_temp_threshold_high': self.rgb_temp_threshold_high,
                'rgb_temp_color_low': self.rgb_temp_color_low,
                'rgb_temp_color_mid': self.rgb_temp_color_mid,
                'rgb_temp_color_high': self.rgb_temp_color_high,
                'auto_hysteresis_c': self.auto_hysteresis_c,
                'auto_debounce_samples': self.auto_debounce_samples,
                'auto_fan_min_toggle_interval_sec': self.auto_fan_min_toggle_interval_sec,
                'auto_pump_min_toggle_interval_sec': self.auto_pump_min_toggle_interval_sec,
                'export_api_enabled': self.export_api_enabled,
                'export_api_port': self.export_api_port,
                'last_device_address': self.last_device_address,
                'last_device_name': self.last_device_name,
                'dingtalk_webhook_enabled': self.dingtalk_webhook_enabled,
                'dingtalk_webhook_url': self.dingtalk_webhook_url,
                'dingtalk_webhook_secret': self.dingtalk_webhook_secret
            }
            config_path = Path(self.CONFIG_FILE)
            config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def set_autostart(self, enable: bool):
        if platform.system() == 'Windows':
            self._set_windows_autostart(enable)
        else:
            self._set_linux_autostart(enable)
        self.auto_start = enable
        self.save()

    def _build_windows_launch_command(self):
        if getattr(sys, 'frozen', False):
            return f'"{Path(sys.executable).resolve()}"'
        return f'"{Path(sys.executable).resolve()}" "{Path(__file__).resolve()}"'

    def _set_windows_run_autostart(self, enable: bool):
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "WaterCoolerManager"
        command = self._build_windows_launch_command()
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        try:
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, command)
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
        finally:
            winreg.CloseKey(key)

    def _set_windows_schtasks_autostart(self, enable: bool):
        task_name = "WaterCoolerManager"
        if enable:
            command = self._build_windows_launch_command()
            subprocess.run(
                [
                    'schtasks', '/Create', '/SC', 'ONLOGON', '/TN', task_name,
                    '/TR', command, '/RL', 'HIGHEST', '/F'
                ],
                check=True,
                capture_output=True,
                text=True,
            )
        else:
            subprocess.run(
                ['schtasks', '/Delete', '/TN', task_name, '/F'],
                check=False,
                capture_output=True,
                text=True,
            )

    def _sync_autostart_if_needed(self):
        try:
            if not self.auto_start:
                return
            if platform.system() == 'Windows':
                self._set_windows_autostart(True)
            else:
                self._set_linux_autostart(True)
        except Exception:
            logging.exception('Failed to sync autostart entry')

    def _set_windows_autostart(self, enable: bool):
        try:
            # 对需要管理员权限的程序，优先使用计划任务（最高权限）以避免登录后自启动失败。
            self._set_windows_schtasks_autostart(enable)
            self._set_windows_run_autostart(False)
            logging.info('Windows autostart updated via Task Scheduler: enable=%s', enable)
        except Exception as exc:
            logging.exception('Task Scheduler autostart update failed, fallback to Run key: %s', exc)
            try:
                self._set_windows_run_autostart(enable)
            except Exception:
                logging.exception('Run key autostart update failed')

    def _set_linux_autostart(self, enable: bool):
        autostart_dir = Path.home() / '.config' / 'autostart'
        autostart_file = autostart_dir / 'watercooler-manager.desktop'
        try:
            if enable:
                autostart_dir.mkdir(parents=True, exist_ok=True)
                exec_path = sys.executable if getattr(sys, 'frozen', False) else str(Path(__file__).resolve())
                autostart_file.write_text(
                    f"[Desktop Entry]\nType=Application\nName=WaterCooler Manager\nExec={exec_path}\nX-GNOME-Autostart-enabled=true\n",
                    encoding='utf-8'
                )
            else:
                if autostart_file.exists():
                    autostart_file.unlink()
        except Exception as e:
            print(f"Error setting autostart: {e}")

COLOR_MAP = {
    '红色':     (255, 0, 0),
    '绿色':   (0, 255, 0),
    '蓝色':    (0, 0, 255),
    '白色':   (255, 255, 255),
    '黄色':  (255, 255, 0),
    '青色':    (0, 255, 255),
    '品红': (255, 0, 255)
}

DLL = Path(__file__).with_name("LibreHardwareMonitorLib.dll")
if not DLL.exists():
    QtWidgets.QMessageBox.critical(None, "Error", f"{DLL.name} not found alongside the script.")
    sys.exit(1)
clr.AddReference(str(DLL))
from LibreHardwareMonitor import Hardware

class Commands:
    FAN  = 0x1B
    PUMP = 0x1C
    RGB  = 0x1E

class PumpVoltage(IntEnum):
    V7  = 0x02
    V8  = 0x03
    V11 = 0x00
    V12 = 0x01  # 协议保留值，界面不再提供

# 水泵电压显示名称
PUMP_VOLTAGE_NAMES = {0x02: "7V", 0x03: "8V", 0x00: "11V", 0x01: "12V（保留协议值）"}

class RGBMode(IntEnum):
    STATIC  = 0x00
    BREATH  = 0x01
    COLORFUL = 0x02  # 彩虹
    BREATHE_COLOR = 0x03  # 呼吸彩虹
    OFF = 0x04  # 关闭

# RGB模式显示名称
RGB_MODE_NAMES = {
    'STATIC': '静态',
    'BREATH': '呼吸',
    'COLORFUL': '彩虹',
    'BREATHE_COLOR': '呼吸彩虹',
    'OFF': '关闭'
}

class NordicUART:
    SERVICE_UUID = '6e400001-b5a3-f393-e0a9-e50e24dcca9e'
    CHAR_TX      = '6e400002-b5a3-f393-e0a9-e50e24dcca9e'

async def scan_devices(models=("LCT21001", "LCT21002", "LCT22002")):
    logging.info('BLE scan start: timeout=6.0, models=%s', ','.join(models))
    devices = await BleakScanner.discover(timeout=6.0)
    logging.info('BLE scan raw result count=%d', len(devices))
    matched = []
    seen = set()
    for d in devices:
        name = (getattr(d, 'name', None) or '').strip()
        addr = (getattr(d, 'address', None) or '').strip()
        if not addr or not name:
            continue
        if not any(m in name.upper() for m in models):
            continue
        key = (name, addr)
        if key in seen:
            continue
        seen.add(key)
        matched.append((name, addr))
    logging.info('BLE scan matched count=%d, devices=%s', len(matched), matched[:8])
    return matched

async def write_fan_mode(client: BleakClient, duty: int):
    safe_duty = clamp_fan_duty(duty)
    packet = bytearray([0xFE, Commands.FAN, 0x01, safe_duty, 0, 0, 0, 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_fan_off(client: BleakClient):
    packet = bytearray([0xFE, Commands.FAN, 0x00, 0, 0, 0, 0, 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_pump_mode(client: BleakClient, voltage: PumpVoltage):
    packet = bytearray([0xFE, Commands.PUMP, 0x01, 100, int(voltage), 0, 0, 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_pump_off(client: BleakClient):
    packet = bytearray([0xFE, Commands.PUMP, 0x00, 0, 0, 0, 0, 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_reset(client: BleakClient):
    packet = bytearray([0xFE, 0x19, 0x00, 1, 0, 0, 0, 0xEF])  # RESET command
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_rgb_mode(client: BleakClient, mode: RGBMode, color: tuple):
    red, green, blue = color
    packet = bytearray([0xFE, Commands.RGB, 0x01, red, green, blue, int(mode), 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

async def write_rgb_off(client: BleakClient):
    packet = bytearray([0xFE, Commands.RGB, 0x00, 0x00, 0x00, 0x00, 0x00, 0xEF])
    await client.write_gatt_char(NordicUART.CHAR_TX, packet)

def get_temperatures():
    try:
        comp = Hardware.Computer()
        comp.IsCpuEnabled = True
        comp.IsGpuEnabled = True
        comp.Open()
        cpu_temp = gpu_temp = None
        
        for hw in comp.Hardware:
            try:
                hw.Update()
                if hw.HardwareType == Hardware.HardwareType.Cpu:
                    # Try different sensor names for CPU temperature
                    for s in hw.Sensors:
                        if s.SensorType == Hardware.SensorType.Temperature:
                            if any(keyword in s.Name for keyword in ["Package", "Core", "CPU"]):
                                if cpu_temp is None or "Package" in s.Name:
                                    cpu_temp = s.Value
                elif hw.HardwareType in (Hardware.HardwareType.GpuNvidia, Hardware.HardwareType.GpuAmd):
                    for s in hw.Sensors:
                        if s.SensorType == Hardware.SensorType.Temperature and "Core" in s.Name:
                            gpu_temp = s.Value
            except Exception:
                pass
        comp.Close()
        return cpu_temp, gpu_temp
    except Exception as e:
        print(f"Error reading temperatures: {e}")
        return None, None

class FanCurveWidget(QtWidgets.QWidget):
    DARK_COLORS = {
        'background': '#111a2c',
        'grid': '#40506b',
        'axis': '#8fa9cb',
        'text': '#dcebff',
        'muted': '#93a9ca',
        'curve': '#57a6ff',
        'point_selected_pen': '#ffffff',
        'point_pen': '#ffd7d7',
        'point_selected_brush': '#ff6b6b',
        'point_brush': '#ff4f6d',
    }
    LIGHT_COLORS = {
        'background': '#f7fbff',
        'grid': '#c9d8ea',
        'axis': '#5e7ea5',
        'text': '#17324d',
        'muted': '#56728f',
        'curve': '#2a78f6',
        'point_selected_pen': '#17324d',
        'point_pen': '#f9b4bf',
        'point_selected_brush': '#ff6b6b',
        'point_brush': '#ff7e93',
    }

    def __init__(self, points):
        super().__init__()
        self.points = sorted(points)
        self.selected = None
        self.dragging = False
        self.selection_changed_callback = None
        self.points_changed_callback = None
        self._left = 58
        self._top = 18
        self._right = 26
        self._bottom = 38
        self._colors = dict(self.DARK_COLORS)
        self.setMinimumHeight(280)
        self.setMouseTracking(True)

    def _notify_selection_changed(self):
        if callable(self.selection_changed_callback):
            self.selection_changed_callback()

    def _notify_points_changed(self):
        if callable(self.points_changed_callback):
            self.points_changed_callback()

    def set_theme_mode(self, theme_mode):
        self._colors = dict(self.DARK_COLORS if theme_mode == 'dark' else self.LIGHT_COLORS)
        self.update()

    def _chart_rect(self):
        return QtCore.QRectF(
            self._left,
            self._top,
            max(120, self.width() - self._left - self._right),
            max(120, self.height() - self._top - self._bottom),
        )

    def _point_to_pos(self, temp, pct):
        rect = self._chart_rect()
        x = rect.left() + (temp - 20) / 80.0 * rect.width()
        y = rect.bottom() - pct / 100.0 * rect.height()
        return QtCore.QPointF(x, y)

    def _pos_to_point(self, x, y):
        rect = self._chart_rect()
        x = min(max(x, rect.left()), rect.right())
        y = min(max(y, rect.top()), rect.bottom())
        temp = 20 + (x - rect.left()) / rect.width() * 80.0
        pct = (rect.bottom() - y) / rect.height() * 100.0
        return int(min(max(temp, 20), 100)), clamp_curve_percent(pct)

    def paintEvent(self, event):
        qp = QtGui.QPainter(self)
        qp.setRenderHint(QtGui.QPainter.Antialiasing)
        c = self._colors
        qp.fillRect(self.rect(), QtGui.QColor(c['background']))

        rect = self._chart_rect()
        grid_pen = QtGui.QPen(QtGui.QColor(c['grid']), 1, QtCore.Qt.DotLine)
        axis_pen = QtGui.QPen(QtGui.QColor(c['axis']), 1.2)
        text_pen = QtGui.QPen(QtGui.QColor(c['text']))
        muted_pen = QtGui.QPen(QtGui.QColor(c['muted']))
        curve_pen = QtGui.QPen(QtGui.QColor(c['curve']), 2.6)
        curve_pen.setCapStyle(QtCore.Qt.RoundCap)
        curve_pen.setJoinStyle(QtCore.Qt.RoundJoin)

        qp.setPen(grid_pen)
        for t in range(20, 101, 10):
            x = self._point_to_pos(t, 0).x()
            qp.drawLine(QtCore.QPointF(x, rect.top()), QtCore.QPointF(x, rect.bottom()))
        for p in range(0, MAX_SAFE_FAN_PERCENT + 1, 10):
            y = self._point_to_pos(20, p).y()
            qp.drawLine(QtCore.QPointF(rect.left(), y), QtCore.QPointF(rect.right(), y))

        qp.setPen(axis_pen)
        qp.drawLine(rect.bottomLeft(), rect.bottomRight())
        qp.drawLine(rect.bottomLeft(), rect.topLeft())

        font = qp.font()
        font.setPointSize(9)
        qp.setFont(font)

        for t in range(20, 101, 20):
            pos = self._point_to_pos(t, 0)
            qp.setPen(axis_pen)
            qp.drawLine(QtCore.QPointF(pos.x(), rect.bottom()), QtCore.QPointF(pos.x(), rect.bottom() + 5))
            qp.setPen(muted_pen)
            qp.drawText(QtCore.QRectF(pos.x() - 22, rect.bottom() + 8, 44, 20), QtCore.Qt.AlignCenter, f'{t}°C')

        for p in (0, 30, 60, MAX_SAFE_FAN_PERCENT):
            pos = self._point_to_pos(20, p)
            qp.setPen(axis_pen)
            qp.drawLine(QtCore.QPointF(rect.left() - 5, pos.y()), QtCore.QPointF(rect.left(), pos.y()))
            qp.setPen(text_pen if p == MAX_SAFE_FAN_PERCENT else muted_pen)
            qp.drawText(QtCore.QRectF(4, pos.y() - 10, self._left - 10, 20), QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, f'{p}%')

        pts = [self._point_to_pos(t, p) for t, p in self.points]
        qp.setPen(curve_pen)
        qp.drawPolyline(QtGui.QPolygonF(pts))

        for i, pt in enumerate(pts):
            is_selected = i == self.selected
            qp.setPen(QtGui.QPen(QtGui.QColor(c['point_selected_pen'] if is_selected else c['point_pen']), 1.6))
            qp.setBrush(QtGui.QColor(c['point_selected_brush'] if is_selected else c['point_brush']))
            qp.drawEllipse(pt, 6, 6)

    def mousePressEvent(self, event):
        x, y = event.x(), event.y()
        self.dragging = False
        for i, (t, p) in enumerate(self.points):
            pos = self._point_to_pos(t, p)
            if abs(pos.x() - x) <= 8 and abs(pos.y() - y) <= 8:
                self.selected = i
                self.dragging = True
                self.update()
                self._notify_selection_changed()
                return
        if self.selected is not None:
            self.selected = None
            self.update()
            self._notify_selection_changed()

    def mouseMoveEvent(self, event):
        if self.selected is None or not self.dragging:
            return
        t, p = self._pos_to_point(event.x(), event.y())

        if self.selected > 0:
            t = max(t, self.points[self.selected - 1][0] + 1)
        if self.selected < len(self.points) - 1:
            t = min(t, self.points[self.selected + 1][0] - 1)

        self.points[self.selected] = (t, p)
        self.update()
        self._notify_points_changed()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        self.update()
        self._notify_selection_changed()

    def interpolate(self, temp):
        pts = sorted(self.points)
        if not pts:
            return 0
        if temp <= pts[0][0]:
            return pts[0][1]
        if temp >= pts[-1][0]:
            return pts[-1][1]
        for i in range(len(pts)-1):
            t0, p0 = pts[i]
            t1, p1 = pts[i+1]
            if t0 <= temp <= t1:
                if t1 == t0:
                    return p1
                return p0 + (p1-p0)*(temp-t0)/(t1-t0)
        return pts[-1][1]


class PumpCurveWidget(QtWidgets.QWidget):
    DARK_COLORS = {
        'background': '#111a2c',
        'grid': '#40506b',
        'axis': '#8fa9cb',
        'text': '#dcebff',
        'muted': '#93a9ca',
        'curve': '#6fd3ff',
        'point_selected_pen': '#ffffff',
        'point_pen': '#ffe8c0',
        'point_selected_brush': '#ffb14f',
        'point_brush': '#ff9a1f',
    }
    LIGHT_COLORS = {
        'background': '#f7fbff',
        'grid': '#c9d8ea',
        'axis': '#5e7ea5',
        'text': '#17324d',
        'muted': '#56728f',
        'curve': '#0096c7',
        'point_selected_pen': '#17324d',
        'point_pen': '#f6c98d',
        'point_selected_brush': '#ffb14f',
        'point_brush': '#ffae42',
    }

    def __init__(self, points):
        super().__init__()
        self.points = normalize_pump_curve_points(points)
        self.selected = None
        self.dragging = False
        self.selection_changed_callback = None
        self.points_changed_callback = None
        self._left = 58
        self._top = 18
        self._right = 26
        self._bottom = 38
        self._colors = dict(self.DARK_COLORS)
        self.setMinimumHeight(280)
        self.setMouseTracking(True)

    def _notify_selection_changed(self):
        if callable(self.selection_changed_callback):
            self.selection_changed_callback()

    def _notify_points_changed(self):
        if callable(self.points_changed_callback):
            self.points_changed_callback()

    def set_theme_mode(self, theme_mode):
        self._colors = dict(self.DARK_COLORS if theme_mode == 'dark' else self.LIGHT_COLORS)
        self.update()

    def _chart_rect(self):
        return QtCore.QRectF(
            self._left,
            self._top,
            max(120, self.width() - self._left - self._right),
            max(120, self.height() - self._top - self._bottom),
        )

    def _level_ratio(self, value):
        value = min(PUMP_CURVE_LEVELS, key=lambda candidate: abs(candidate - int(value)))
        return PUMP_CURVE_LEVELS.index(value) / float(len(PUMP_CURVE_LEVELS) - 1)

    def _point_to_pos(self, temp, value):
        rect = self._chart_rect()
        x = rect.left() + (temp - 20) / 80.0 * rect.width()
        y = rect.bottom() - self._level_ratio(value) * rect.height()
        return QtCore.QPointF(x, y)

    def _pos_to_point(self, x, y):
        rect = self._chart_rect()
        x = min(max(x, rect.left()), rect.right())
        y = min(max(y, rect.top()), rect.bottom())
        temp = 20 + (x - rect.left()) / rect.width() * 80.0
        ratio = (rect.bottom() - y) / rect.height()
        idx = min(range(len(PUMP_CURVE_LEVELS)), key=lambda i: abs(i / float(len(PUMP_CURVE_LEVELS) - 1) - ratio))
        return int(min(max(temp, 20), 100)), PUMP_CURVE_LEVELS[idx]

    def paintEvent(self, event):
        qp = QtGui.QPainter(self)
        qp.setRenderHint(QtGui.QPainter.Antialiasing)
        c = self._colors
        qp.fillRect(self.rect(), QtGui.QColor(c['background']))

        rect = self._chart_rect()
        grid_pen = QtGui.QPen(QtGui.QColor(c['grid']), 1, QtCore.Qt.DotLine)
        axis_pen = QtGui.QPen(QtGui.QColor(c['axis']), 1.2)
        text_pen = QtGui.QPen(QtGui.QColor(c['text']))
        muted_pen = QtGui.QPen(QtGui.QColor(c['muted']))
        curve_pen = QtGui.QPen(QtGui.QColor(c['curve']), 2.6)
        curve_pen.setCapStyle(QtCore.Qt.RoundCap)
        curve_pen.setJoinStyle(QtCore.Qt.RoundJoin)

        qp.setPen(grid_pen)
        for t in range(20, 101, 10):
            x = self._point_to_pos(t, PUMP_CURVE_LEVELS[0]).x()
            qp.drawLine(QtCore.QPointF(x, rect.top()), QtCore.QPointF(x, rect.bottom()))
        for value in PUMP_CURVE_LEVELS:
            y = self._point_to_pos(20, value).y()
            qp.drawLine(QtCore.QPointF(rect.left(), y), QtCore.QPointF(rect.right(), y))

        qp.setPen(axis_pen)
        qp.drawLine(rect.bottomLeft(), rect.bottomRight())
        qp.drawLine(rect.bottomLeft(), rect.topLeft())

        font = qp.font()
        font.setPointSize(9)
        qp.setFont(font)

        for t in range(20, 101, 20):
            pos = self._point_to_pos(t, PUMP_CURVE_LEVELS[0])
            qp.setPen(axis_pen)
            qp.drawLine(QtCore.QPointF(pos.x(), rect.bottom()), QtCore.QPointF(pos.x(), rect.bottom() + 5))
            qp.setPen(muted_pen)
            qp.drawText(QtCore.QRectF(pos.x() - 22, rect.bottom() + 8, 44, 20), QtCore.Qt.AlignCenter, f'{t}°C')

        for value in PUMP_CURVE_LEVELS:
            pos = self._point_to_pos(20, value)
            qp.setPen(axis_pen)
            qp.drawLine(QtCore.QPointF(rect.left() - 5, pos.y()), QtCore.QPointF(rect.left(), pos.y()))
            qp.setPen(text_pen if value == PUMP_CURVE_LEVELS[-1] else muted_pen)
            qp.drawText(QtCore.QRectF(4, pos.y() - 10, self._left - 10, 20), QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, pump_curve_value_to_text(value))

        pts = [self._point_to_pos(t, value) for t, value in self.points]
        qp.setPen(curve_pen)
        qp.drawPolyline(QtGui.QPolygonF(pts))

        for i, pt in enumerate(pts):
            is_selected = i == self.selected
            qp.setPen(QtGui.QPen(QtGui.QColor(c['point_selected_pen'] if is_selected else c['point_pen']), 1.6))
            qp.setBrush(QtGui.QColor(c['point_selected_brush'] if is_selected else c['point_brush']))
            qp.drawEllipse(pt, 6, 6)

    def mousePressEvent(self, event):
        x, y = event.x(), event.y()
        self.dragging = False
        for i, (t, value) in enumerate(self.points):
            pos = self._point_to_pos(t, value)
            if abs(pos.x() - x) <= 8 and abs(pos.y() - y) <= 8:
                self.selected = i
                self.dragging = True
                self.update()
                self._notify_selection_changed()
                return
        if self.selected is not None:
            self.selected = None
            self.update()
            self._notify_selection_changed()

    def mouseMoveEvent(self, event):
        if self.selected is None or not self.dragging:
            return
        temp, value = self._pos_to_point(event.x(), event.y())
        if self.selected > 0:
            temp = max(temp, self.points[self.selected - 1][0] + 1)
        if self.selected < len(self.points) - 1:
            temp = min(temp, self.points[self.selected + 1][0] - 1)
        self.points[self.selected] = (temp, value)
        self.update()
        self._notify_points_changed()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        self.update()
        self._notify_selection_changed()

    def interpolate(self, temp):
        pts = sorted(self.points)
        if not pts:
            return 0
        if temp <= pts[0][0]:
            return pts[0][1]
        if temp >= pts[-1][0]:
            return pts[-1][1]
        for i in range(len(pts)-1):
            t0, v0 = pts[i]
            t1, v1 = pts[i+1]
            if t0 <= temp <= t1:
                if t1 == t0:
                    return v1
                ratio = (temp - t0) / (t1 - t0)
                numeric = v0 + (v1 - v0) * ratio
                return min(PUMP_CURVE_LEVELS, key=lambda candidate: abs(candidate - numeric))
        return pts[-1][1]


class MainWindow(QtWidgets.QMainWindow):
    dingtalk_test_result = QtCore.pyqtSignal(bool, str)

    UPDATE_INTERVALS = [
        (0.5, "0.5 秒"),
        (1.0, "1 秒"),
        (2.0, "2 秒"),
        (3.0, "3 秒"),
        (5.0, "5 秒"),
        (10.0, "10 秒"),
    ]
    DEFAULT_INTERVAL_SEC = DEFAULT_UPDATE_INTERVAL_SEC

    def __init__(self):
        super().__init__()
        self.dingtalk_test_result.connect(self._on_dingtalk_test_result)
        self.setWindowTitle("水冷管理器")
        self.setMinimumWidth(700)
        self.setAttribute(QtCore.Qt.WA_QuitOnClose, False)
        icon_dir = Path(__file__).parent / 'icons'
        ico = icon_dir / 'water_drop.ico'
        png = icon_dir / 'water_drop.png'
        icon_file = str(ico if ico.exists() else png)
        icon = QtGui.QIcon(icon_file)
        self.tray_icon = QtWidgets.QSystemTrayIcon(icon, self)
        self.setWindowIcon(icon)
        tray_menu = QtWidgets.QMenu()
        tray_menu.addAction(UI['tray_show'], self.show_window)
        tray_menu.addAction(UI['tray_exit'], self.exit_app)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()
        self.client = None
        self.fan_curve_points = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
        self.pump_curve_points = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
        self.auto_mode_active = False
        self.last_cpu_temp = None
        self.last_gpu_temp = None
        self.is_scanning = False
        self.is_connecting = False
        self.is_disconnecting = False
        self.settings = Settings()
        self.pump_runtime_on = False
        self.pump_runtime_voltage = None
        self._is_exiting = False
        self._restore_maximized = True
        self._last_temp_rgb_bucket = None
        self._syncing_ui = False
        self._control_temp_history = deque(maxlen=self.settings.auto_debounce_samples)
        self._auto_applied_fan_percent = None
        self._auto_applied_pump_value = None
        self._last_fan_toggle_ts = 0.0
        self._last_pump_toggle_ts = 0.0
        self._manual_prompt_suspend = False
        self._resolved_theme_mode = self._get_effective_theme_mode()
        self.export_api_state = ExportApiState()
        self.export_api_server = None
        self._temperature_update_in_progress = False
        self._disconnect_callback_scheduled = False
        self._auto_reconnect_pending = False
        self._build_ui()
        self.sync_ui_from_settings()
        self._apply_export_api_settings(save=False)
        self.temp_timer = QtCore.QTimer(self)
        self.temp_timer.timeout.connect(self.update_temperatures)
        self.set_update_interval(self.settings.update_interval_sec, save=False)
        self.temp_timer.start()
        self._refresh_export_api_state()

    def _detect_system_theme_mode(self):
        if platform.system() == 'Windows':
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                apps_use_light_theme = int(winreg.QueryValueEx(key, "AppsUseLightTheme")[0])
                winreg.CloseKey(key)
                return 'light' if apps_use_light_theme else 'dark'
            except Exception:
                pass
        app = QtWidgets.QApplication.instance()
        if app is not None:
            color = app.palette().window().color()
            brightness = (color.red() * 299 + color.green() * 587 + color.blue() * 114) / 1000
            return 'dark' if brightness < 144 else 'light'
        return 'dark'

    def _get_effective_theme_mode(self):
        mode = getattr(self.settings, 'theme_mode', 'system') if hasattr(self, 'settings') else 'system'
        return self._detect_system_theme_mode() if mode == 'system' else mode

    def _theme_tokens(self):
        mode = self._resolved_theme_mode
        if mode == 'light':
            return {
                'main_bg': '#edf3fb', 'text': '#1e3550', 'card_bg': '#ffffff', 'card_border': '#cddbeb',
                'hero_g0': '#f8fbff', 'hero_g1': '#eef5ff', 'hero_border': '#c8daf6', 'hero_title': '#16324d',
                'hero_subtitle': '#557391', 'muted': '#617b96', 'badge_blue_bg': 'rgba(42, 120, 246, 0.10)',
                'badge_blue_text': '#2a78f6', 'badge_blue_border': 'rgba(42, 120, 246, 0.20)',
                'badge_green_bg': 'rgba(39, 174, 96, 0.12)', 'badge_green_text': '#1f8a4d',
                'badge_green_border': 'rgba(39, 174, 96, 0.22)', 'badge_red_bg': 'rgba(229, 83, 83, 0.10)',
                'badge_red_text': '#c03b3b', 'badge_red_border': 'rgba(229, 83, 83, 0.18)', 'info_title': '#5f7b96',
                'info_value': '#16324d', 'section_title': '#17324d', 'section_subtitle': '#68809a',
                'pill_bg': '#f4f8fd', 'pill_border': '#cad7e7', 'pill_text': '#274868', 'status_tag_bg': 'rgba(42, 120, 246, 0.10)',
                'status_tag_border': 'rgba(42, 120, 246, 0.18)', 'status_tag_text': '#2a78f6', 'status_text': '#2f4c68',
                'combo_bg': '#ffffff', 'combo_text': '#17324d', 'combo_border': '#c7d6e8', 'combo_popup_bg': '#ffffff',
                'combo_popup_border': '#c7d6e8', 'combo_popup_select_bg': '#eaf2fe', 'button_bg': '#2c6cf3',
                'button_hover': '#3b7afd', 'button_pressed': '#1f57cf', 'button_disabled_bg': '#dce5ef',
                'button_disabled_text': '#7c90a5', 'ghost_bg': '#f5f9ff', 'ghost_border': '#c7d7ea', 'ghost_text': '#1f4265',
                'ghost_hover': '#edf4ff', 'check_border': '#9eb4cb', 'check_bg': '#ffffff', 'slider_bg': '#dbe5f0',
                'group_border': '#cfdaea', 'group_title': '#17324d', 'scroll_bg': '#e9f0f8', 'scroll_handle': '#bac8d9',
            }
        return {
            'main_bg': '#0b1220', 'text': '#e5eefc', 'card_bg': '#131d31', 'card_border': '#243149',
            'hero_g0': '#18253d', 'hero_g1': '#102436', 'hero_border': '#2f4a73', 'hero_title': '#f8fbff',
            'hero_subtitle': '#a8bddb', 'muted': '#9db1ce', 'badge_blue_bg': 'rgba(74, 155, 255, 0.16)',
            'badge_blue_text': '#9bc7ff', 'badge_blue_border': 'rgba(100, 175, 255, 0.28)',
            'badge_green_bg': 'rgba(57, 211, 83, 0.18)', 'badge_green_text': '#78f28f',
            'badge_green_border': 'rgba(92, 230, 116, 0.35)', 'badge_red_bg': 'rgba(255, 97, 97, 0.14)',
            'badge_red_text': '#ff9b9b', 'badge_red_border': 'rgba(255, 135, 135, 0.26)', 'info_title': '#8fa9cb',
            'info_value': '#f4f8ff', 'section_title': '#f6fbff', 'section_subtitle': '#8fa4c5',
            'pill_bg': '#0f1728', 'pill_border': '#2a3954', 'pill_text': '#cfe3ff', 'status_tag_bg': 'rgba(77, 163, 255, 0.16)',
            'status_tag_border': 'rgba(90, 175, 255, 0.28)', 'status_tag_text': '#9bc7ff', 'status_text': '#dce8fa',
            'combo_bg': '#0d1524', 'combo_text': '#eef5ff', 'combo_border': '#2c3a53', 'combo_popup_bg': '#10192b',
            'combo_popup_border': '#304261', 'combo_popup_select_bg': '#1e3353', 'button_bg': '#2c6cf3',
            'button_hover': '#3b7afd', 'button_pressed': '#1f57cf', 'button_disabled_bg': '#263246',
            'button_disabled_text': '#7990b1', 'ghost_bg': '#182236', 'ghost_border': '#334562', 'ghost_text': '#dbe8fb',
            'ghost_hover': '#1b2a43', 'check_border': '#436087', 'check_bg': '#0c1422', 'slider_bg': '#0c1422',
            'group_border': '#2a3850', 'group_title': '#ddebff', 'scroll_bg': '#0f1728', 'scroll_handle': '#314766',
        }

    def _apply_theme_to_curve_widgets(self):
        if hasattr(self, 'curve_widget'):
            self.curve_widget.set_theme_mode(self._resolved_theme_mode)
        if hasattr(self, 'pump_curve_widget'):
            self.pump_curve_widget.set_theme_mode(self._resolved_theme_mode)

    def _apply_styles(self):
        c = self._theme_tokens()
        self.setStyleSheet(f"""
            QMainWindow {{ background: {c['main_bg']}; }}
            QWidget {{ color: {c['text']}; font-size: 13px; }}
            QFrame#heroCard, QFrame#footerCard, QFrame#toolbarCard, QFrame#panelCard, QFrame#infoCard {{
                background: {c['card_bg']}; border: 1px solid {c['card_border']}; border-radius: 18px;
            }}
            QStackedWidget, QScrollArea, QScrollArea > QWidget > QWidget {{ background: transparent; border: none; }}
            QFrame#heroCard {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {c['hero_g0']}, stop:1 {c['hero_g1']});
                border: 1px solid {c['hero_border']};
            }}
            QLabel#heroTitle {{ font-size: 28px; font-weight: 700; color: {c['hero_title']}; }}
            QLabel#heroSubtitle {{ font-size: 14px; color: {c['hero_subtitle']}; }}
            QLabel#heroMeta, QLabel#mutedText {{ color: {c['muted']}; }}
            QLabel#statusBadgeConnected, QLabel#statusBadgeDisconnected, QLabel#secondaryBadge {{
                padding: 7px 14px; border-radius: 14px; font-weight: 600;
            }}
            QLabel#statusBadgeConnected {{ background: {c['badge_green_bg']}; color: {c['badge_green_text']}; border: 1px solid {c['badge_green_border']}; }}
            QLabel#statusBadgeDisconnected {{ background: {c['badge_red_bg']}; color: {c['badge_red_text']}; border: 1px solid {c['badge_red_border']}; }}
            QLabel#secondaryBadge {{ background: {c['badge_blue_bg']}; color: {c['badge_blue_text']}; border: 1px solid {c['badge_blue_border']}; }}
            QLabel#infoTitle {{ color: {c['info_title']}; font-size: 12px; }}
            QLabel#infoValue {{ font-size: 22px; font-weight: 700; color: {c['info_value']}; }}
            QLabel#sectionTitle {{ font-size: 16px; font-weight: 700; color: {c['section_title']}; }}
            QLabel#sectionSubtitle, QLabel#hintText {{ color: {c['section_subtitle']}; font-size: 12px; }}
            QLabel#sectionMiniTitle {{ font-size: 13px; font-weight: 600; color: {c['section_title']}; }}
            QLabel#valuePill {{ background: {c['pill_bg']}; border: 1px solid {c['pill_border']}; border-radius: 12px; padding: 6px 10px; font-weight: 600; color: {c['pill_text']}; }}
            QLabel#statusTag {{ background: {c['status_tag_bg']}; border: 1px solid {c['status_tag_border']}; border-radius: 12px; padding: 6px 10px; font-weight: 600; color: {c['status_tag_text']}; }}
            QLabel#statusText {{ color: {c['status_text']}; }}
            QLabel#metricHighlight {{ color: {c['info_value']}; font-size: 14px; font-weight: 600; }}
            QComboBox, QSpinBox, QDoubleSpinBox {{ color: {c['combo_text']}; background: {c['combo_bg']}; border: 1px solid {c['combo_border']}; border-radius: 10px; padding: 7px 10px; min-height: 20px; }}
            QComboBox::drop-down {{ border: none; width: 24px; }}
            QComboBox QAbstractItemView {{ background: {c['combo_popup_bg']}; color: {c['combo_text']}; border: 1px solid {c['combo_popup_border']}; selection-background-color: {c['combo_popup_select_bg']}; selection-color: {c['combo_text']}; }}
            QPushButton {{ background: {c['button_bg']}; color: white; border: none; border-radius: 12px; padding: 10px 18px; font-weight: 600; }}
            QPushButton:hover {{ background: {c['button_hover']}; }}
            QPushButton:pressed {{ background: {c['button_pressed']}; }}
            QPushButton:disabled {{ background: {c['button_disabled_bg']}; color: {c['button_disabled_text']}; }}
            QPushButton#ghostButton {{ background: {c['ghost_bg']}; border: 1px solid {c['ghost_border']}; color: {c['ghost_text']}; }}
            QPushButton#ghostButton:hover {{ background: {c['ghost_hover']}; }}
            QCheckBox {{ spacing: 8px; }}
            QCheckBox::indicator {{ width: 18px; height: 18px; border-radius: 6px; border: 1px solid {c['check_border']}; background: {c['check_bg']}; }}
            QCheckBox::indicator:checked {{ background: {c['button_bg']}; border: 1px solid {c['button_bg']}; }}
            QSlider::groove:horizontal {{ height: 8px; background: {c['slider_bg']}; border-radius: 4px; }}
            QSlider::sub-page:horizontal {{ background: {c['button_bg']}; border-radius: 4px; }}
            QSlider::handle:horizontal {{ background: #ffffff; border: 2px solid {c['button_bg']}; width: 18px; margin: -6px 0; border-radius: 9px; }}
            QGroupBox {{ margin-top: 14px; border: 1px solid {c['group_border']}; border-radius: 14px; padding-top: 12px; font-weight: 600; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 6px; color: {c['group_title']}; }}
            QScrollBar:vertical {{ background: {c['scroll_bg']}; width: 10px; margin: 4px; border-radius: 5px; }}
            QScrollBar::handle:vertical {{ background: {c['scroll_handle']}; min-height: 24px; border-radius: 5px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; height: 0px; }}
        """)
        self._apply_theme_to_curve_widgets()

    def _apply_theme(self, save=False):
        self._resolved_theme_mode = self._get_effective_theme_mode()
        self._apply_styles()
        if hasattr(self, 'connection_badge'):
            self.connection_badge.style().unpolish(self.connection_badge)
            self.connection_badge.style().polish(self.connection_badge)
        if save:
            self.settings.save()

    def _create_info_card(self, parent_layout, title, value):
        card = QtWidgets.QFrame()
        card.setObjectName("infoCard")
        card_layout = QtWidgets.QVBoxLayout(card)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(6)
        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("infoTitle")
        value_label = QtWidgets.QLabel(value)
        value_label.setObjectName("infoValue")
        card_layout.addWidget(title_label)
        card_layout.addWidget(value_label)
        card_layout.addStretch()
        parent_layout.addWidget(card, 1)
        return value_label

    def _create_panel(self, title, subtitle=None):
        frame = QtWidgets.QFrame()
        frame.setObjectName("panelCard")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)
        header = QtWidgets.QVBoxLayout()
        header.setSpacing(3)
        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("sectionTitle")
        header.addWidget(title_label)
        if subtitle:
            subtitle_label = QtWidgets.QLabel(subtitle)
            subtitle_label.setObjectName("sectionSubtitle")
            subtitle_label.setWordWrap(True)
            header.addWidget(subtitle_label)
        layout.addLayout(header)
        return frame, layout

    def _wrap_page_in_scroll(self, widget):
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll.setWidget(widget)
        return scroll

    def _build_slider_scale(self, labels, values=None):
        wrapper = QtWidgets.QVBoxLayout()
        wrapper.setSpacing(6)
        labels_row = QtWidgets.QHBoxLayout()
        labels_row.setSpacing(8)
        for label in labels:
            item = QtWidgets.QLabel(label)
            item.setAlignment(QtCore.Qt.AlignCenter)
            item.setObjectName("hintText")
            labels_row.addWidget(item)
        wrapper.addLayout(labels_row)
        if values:
            values_row = QtWidgets.QHBoxLayout()
            values_row.setSpacing(8)
            for value in values:
                item = QtWidgets.QLabel(value)
                item.setAlignment(QtCore.Qt.AlignCenter)
                item.setObjectName("mutedText")
                values_row.addWidget(item)
            wrapper.addLayout(values_row)
        return wrapper

    def _build_manual_page(self):
        manual = QtWidgets.QWidget()
        page_layout = QtWidgets.QVBoxLayout(manual)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(14)

        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(14)

        connect_card, connect_layout = self._create_panel("设备连接", "扫描到设备后即可连接；支持启动自动连接。")
        device_row = QtWidgets.QHBoxLayout()
        device_row.setSpacing(10)
        self.device_combo = QtWidgets.QComboBox()
        self.device_combo.setEnabled(False)
        self.rescan_btn = QtWidgets.QPushButton("重新扫描")
        self.rescan_btn.setObjectName("ghostButton")
        self.rescan_btn.setToolTip("立即重新扫描可连接的水冷设备")
        self.connect_btn = QtWidgets.QPushButton(UI['btn_connect'])
        self.connect_btn.setEnabled(False)
        device_row.addWidget(self.device_combo, 1)
        device_row.addWidget(self.rescan_btn)
        device_row.addWidget(self.connect_btn)
        connect_layout.addLayout(device_row)
        self.device_tip_label = QtWidgets.QLabel("当前未发现可连接设备")
        self.device_tip_label.setObjectName("mutedText")
        self.device_tip_label.setWordWrap(True)
        connect_layout.addWidget(self.device_tip_label)
        connect_card.setMinimumHeight(168)
        top_row.addWidget(connect_card, 3)

        quick_card, quick_layout = self._create_panel("快捷操作", "手动模式下可一键切换预设；拖动风扇或水泵后，需点击应用才会真正下发到设备。")
        preset_grid = QtWidgets.QGridLayout()
        preset_grid.setHorizontalSpacing(10)
        preset_grid.setVerticalSpacing(10)
        self.preset_silent_btn = QtWidgets.QPushButton("静音模式 30% / 7V")
        self.preset_silent_btn.setObjectName("ghostButton")
        self.preset_silent_btn.setEnabled(False)
        self.preset_balanced_btn = QtWidgets.QPushButton("平衡模式 60% / 8V")
        self.preset_balanced_btn.setObjectName("ghostButton")
        self.preset_balanced_btn.setEnabled(False)
        self.preset_performance_btn = QtWidgets.QPushButton("高效模式 90% / 11V")
        self.preset_performance_btn.setObjectName("ghostButton")
        self.preset_performance_btn.setEnabled(False)
        preset_grid.addWidget(self.preset_silent_btn, 0, 0)
        preset_grid.addWidget(self.preset_balanced_btn, 0, 1)
        preset_grid.addWidget(self.preset_performance_btn, 1, 0, 1, 2)
        quick_layout.addLayout(preset_grid)

        button_grid = QtWidgets.QGridLayout()
        button_grid.setHorizontalSpacing(10)
        button_grid.setVerticalSpacing(10)
        self.apply_manual_btn = QtWidgets.QPushButton(UI['btn_apply_manual'])
        self.apply_manual_btn.setEnabled(False)
        self.apply_rgb_btn = QtWidgets.QPushButton(UI['btn_apply_rgb'])
        self.apply_rgb_btn.setEnabled(False)
        self.apply_rgb_btn.setObjectName("ghostButton")
        self.apply_all_btn = QtWidgets.QPushButton(UI['btn_apply_all'])
        self.apply_all_btn.setEnabled(False)
        button_grid.addWidget(self.apply_manual_btn, 0, 0)
        button_grid.addWidget(self.apply_all_btn, 0, 1)
        quick_layout.addLayout(button_grid)

        self.manual_apply_hint_label = QtWidgets.QLabel("拖动风扇或水泵后，需要点击“应用”按钮才会生效。")
        self.manual_apply_hint_label.setObjectName("hintText")
        self.manual_apply_hint_label.setWordWrap(True)
        quick_layout.addWidget(self.manual_apply_hint_label)
        quick_card.setMinimumHeight(230)
        top_row.addWidget(quick_card, 2)
        page_layout.addLayout(top_row)

        control_row = QtWidgets.QHBoxLayout()
        control_row.setSpacing(14)

        left_col = QtWidgets.QVBoxLayout()
        left_col.setSpacing(14)
        fan_card, fan_layout = self._create_panel("风扇控制", "手动模式下直接按百分比控制；拖动精度为 1%，最高按上游安全限制钳制到 90%。")
        fan_header = QtWidgets.QHBoxLayout()
        fan_header.addWidget(QtWidgets.QLabel(UI['label_fan']))
        fan_header.addStretch()
        self.fan_value_pill = QtWidgets.QLabel("50%")
        self.fan_value_pill.setObjectName("valuePill")
        fan_header.addWidget(self.fan_value_pill)
        fan_layout.addLayout(fan_header)
        self.fan_slider = NoWheelSlider(QtCore.Qt.Horizontal)
        self.fan_slider.setRange(0, MAX_SAFE_FAN_PERCENT)
        self.fan_slider.setSingleStep(1)
        self.fan_slider.setPageStep(5)
        self.fan_slider.setTickInterval(10)
        self.fan_slider.setTickPosition(QtWidgets.QSlider.TicksBelow)
        fan_layout.addWidget(self.fan_slider)
        fan_layout.addLayout(self._build_slider_scale(["0%", "30%", "60%", "90%"]))
        fan_card.setMinimumHeight(188)
        left_col.addWidget(fan_card)

        pump_card, pump_layout = self._create_panel("水泵控制", "手动模式下直接设置固定电压；为保护水泵，最高限制为 11V；自动模式下可使用水泵曲线。")
        pump_header = QtWidgets.QHBoxLayout()
        pump_header.addWidget(QtWidgets.QLabel(UI['label_pump']))
        pump_header.addStretch()
        self.pump_value_pill = QtWidgets.QLabel("7V")
        self.pump_value_pill.setObjectName("valuePill")
        pump_header.addWidget(self.pump_value_pill)
        pump_layout.addLayout(pump_header)
        self.pump_slider = NoWheelSlider(QtCore.Qt.Horizontal)
        self.pump_slider.setRange(0, 3)
        self.pump_slider.setTickInterval(1)
        self.pump_slider.setTickPosition(QtWidgets.QSlider.TicksBelow)
        pump_layout.addWidget(self.pump_slider)
        pump_layout.addLayout(self._build_slider_scale(["关闭", "7V", "8V", "11V"]))
        manual_hint = QtWidgets.QLabel("提示：自动模式会根据温度曲线同时调节风扇与水泵；手动模式不会保留托管控制。")
        manual_hint.setObjectName("hintText")
        manual_hint.setWordWrap(True)
        pump_layout.addWidget(manual_hint)
        pump_card.setMinimumHeight(188)
        left_col.addWidget(pump_card)
        control_row.addLayout(left_col, 3)

        right_col = QtWidgets.QVBoxLayout()
        right_col.setSpacing(14)
        rgb_card, rgb_layout = self._create_panel("RGB 灯效", "支持常规灯效，也支持按温度自动切换颜色。")
        rgb_header = QtWidgets.QHBoxLayout()
        rgb_header.addWidget(QtWidgets.QLabel(UI['label_rgb']))
        rgb_header.addStretch()
        self.rgb_value_pill = QtWidgets.QLabel("静态 · 红色")
        self.rgb_value_pill.setObjectName("valuePill")
        rgb_header.addWidget(self.rgb_value_pill)
        rgb_layout.addLayout(rgb_header)
        rgb_controls = QtWidgets.QHBoxLayout()
        rgb_controls.setSpacing(10)
        self.rgb_mode = NoWheelComboBox()
        for n, m in RGBMode.__members__.items():
            display_name = RGB_MODE_NAMES.get(n, n)
            self.rgb_mode.addItem(display_name, m)
        self.rgb_color = NoWheelComboBox()
        for n, c in COLOR_MAP.items():
            self.rgb_color.addItem(n, c)
        self.apply_rgb_btn.setMinimumWidth(110)
        rgb_controls.addWidget(self.rgb_mode, 2)
        rgb_controls.addWidget(self.rgb_color, 1)
        rgb_controls.addWidget(self.apply_rgb_btn)
        rgb_layout.addLayout(rgb_controls)

        self.rgb_temp_enabled_checkbox = QtWidgets.QCheckBox("启用温控 RGB")
        rgb_layout.addWidget(self.rgb_temp_enabled_checkbox)

        temp_mode_row = QtWidgets.QHBoxLayout()
        temp_mode_row.setSpacing(10)
        temp_mode_row.addWidget(QtWidgets.QLabel("联动模式"))
        self.rgb_temp_mode_combo = NoWheelComboBox()
        self.rgb_temp_mode_combo.addItem("静态", RGBMode.STATIC)
        self.rgb_temp_mode_combo.addItem("呼吸", RGBMode.BREATH)
        temp_mode_row.addWidget(self.rgb_temp_mode_combo, 1)
        rgb_layout.addLayout(temp_mode_row)

        threshold_row = QtWidgets.QGridLayout()
        threshold_row.setHorizontalSpacing(10)
        threshold_row.setVerticalSpacing(8)
        threshold_row.addWidget(QtWidgets.QLabel("低温上限"), 0, 0)
        self.rgb_temp_low_spin = NoWheelSpinBox()
        self.rgb_temp_low_spin.setRange(20, 95)
        self.rgb_temp_low_spin.setSuffix(" °C")
        threshold_row.addWidget(self.rgb_temp_low_spin, 0, 1)
        threshold_row.addWidget(QtWidgets.QLabel("中温上限"), 0, 2)
        self.rgb_temp_high_spin = NoWheelSpinBox()
        self.rgb_temp_high_spin.setRange(21, 100)
        self.rgb_temp_high_spin.setSuffix(" °C")
        threshold_row.addWidget(self.rgb_temp_high_spin, 0, 3)
        rgb_layout.addLayout(threshold_row)

        color_grid = QtWidgets.QGridLayout()
        color_grid.setHorizontalSpacing(10)
        color_grid.setVerticalSpacing(8)
        color_grid.addWidget(QtWidgets.QLabel("低温颜色"), 0, 0)
        self.rgb_temp_low_color = NoWheelComboBox()
        color_grid.addWidget(self.rgb_temp_low_color, 0, 1)
        color_grid.addWidget(QtWidgets.QLabel("中温颜色"), 1, 0)
        self.rgb_temp_mid_color = NoWheelComboBox()
        color_grid.addWidget(self.rgb_temp_mid_color, 1, 1)
        color_grid.addWidget(QtWidgets.QLabel("高温颜色"), 2, 0)
        self.rgb_temp_high_color = NoWheelComboBox()
        color_grid.addWidget(self.rgb_temp_high_color, 2, 1)
        for combo in (self.rgb_temp_low_color, self.rgb_temp_mid_color, self.rgb_temp_high_color):
            for n, c in COLOR_MAP.items():
                combo.addItem(n, c)
        rgb_layout.addLayout(color_grid)

        self.rgb_help_label = QtWidgets.QLabel("提示：开启温控 RGB 后，会按当前温度自动切换低/中/高三段颜色。")
        self.rgb_help_label.setObjectName("hintText")
        self.rgb_help_label.setWordWrap(True)
        rgb_layout.addWidget(self.rgb_help_label)
        rgb_card.setMinimumHeight(360)
        right_col.addWidget(rgb_card)

        overview_card, overview_layout = self._create_panel("当前配置预览", "这里展示将要下发到设备的主要设置，便于快速确认。")
        preview_form = QtWidgets.QFormLayout()
        preview_form.setLabelAlignment(QtCore.Qt.AlignLeft)
        preview_form.setFormAlignment(QtCore.Qt.AlignTop)
        preview_form.setHorizontalSpacing(16)
        preview_form.setVerticalSpacing(10)
        self.preview_mode_label = QtWidgets.QLabel("手动模式")
        self.preview_mode_label.setObjectName("valuePill")
        self.preview_fan_label = QtWidgets.QLabel("30%")
        self.preview_fan_label.setObjectName("valuePill")
        self.preview_pump_label = QtWidgets.QLabel("7V")
        self.preview_pump_label.setObjectName("valuePill")
        self.preview_rgb_label = QtWidgets.QLabel("温控 / 静态 / 低温 白色")
        self.preview_rgb_label.setObjectName("valuePill")
        preview_form.addRow("界面模式", self.preview_mode_label)
        preview_form.addRow("风扇设置", self.preview_fan_label)
        preview_form.addRow("水泵设置", self.preview_pump_label)
        preview_form.addRow("RGB 设置", self.preview_rgb_label)
        overview_layout.addLayout(preview_form)
        overview_card.setMinimumHeight(188)
        right_col.addWidget(overview_card)
        right_col.addStretch()
        control_row.addLayout(right_col, 2)

        page_layout.addLayout(control_row)
        return manual

    def _build_auto_page(self):
        auto = QtWidgets.QWidget()
        page_layout = QtWidgets.QVBoxLayout(auto)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(14)

        top_row = QtWidgets.QHBoxLayout()
        top_row.setSpacing(14)

        fan_card, fan_layout = self._create_panel("自动风扇曲线", "拖动控制点定义 CPU 温度与风扇百分比映射；默认 4 个点，支持增加和删除坐标点。")
        self.curve_widget = FanCurveWidget(self.fan_curve_points)
        self.curve_widget.selection_changed_callback = self._update_curve_editor_buttons
        self.curve_widget.points_changed_callback = self._on_curve_points_edited
        fan_layout.addWidget(self.curve_widget, 1)
        fan_actions = QtWidgets.QHBoxLayout()
        self.apply_curve_btn = QtWidgets.QPushButton(UI['btn_apply_curve'])
        self.apply_curve_btn.setEnabled(False)
        self.add_fan_curve_btn = QtWidgets.QPushButton("增加坐标点")
        self.add_fan_curve_btn.setObjectName("ghostButton")
        self.remove_fan_curve_btn = QtWidgets.QPushButton("删除选中点")
        self.remove_fan_curve_btn.setObjectName("ghostButton")
        self.reset_curve_btn = QtWidgets.QPushButton("重置风扇曲线")
        self.reset_curve_btn.setObjectName("ghostButton")
        fan_actions.addWidget(self.apply_curve_btn)
        fan_actions.addWidget(self.add_fan_curve_btn)
        fan_actions.addWidget(self.remove_fan_curve_btn)
        fan_actions.addWidget(self.reset_curve_btn)
        fan_actions.addStretch()
        fan_layout.addLayout(fan_actions)
        self.fan_curve_preview_label = QtWidgets.QLabel("当前温度：-- °C，预计风扇：-- %")
        self.fan_curve_preview_label.setObjectName("valuePill")
        fan_layout.addWidget(self.fan_curve_preview_label)
        fan_card.setMinimumHeight(450)
        top_row.addWidget(fan_card, 3)

        pump_card, pump_layout = self._create_panel("自动水泵曲线", "拖动控制点定义 CPU 温度与水泵电压映射；默认 4 个点，支持增加和删除坐标点。")
        self.pump_curve_widget = PumpCurveWidget(self.pump_curve_points)
        self.pump_curve_widget.selection_changed_callback = self._update_curve_editor_buttons
        self.pump_curve_widget.points_changed_callback = self._on_curve_points_edited
        pump_layout.addWidget(self.pump_curve_widget, 1)
        pump_actions = QtWidgets.QHBoxLayout()
        self.add_pump_curve_btn = QtWidgets.QPushButton("增加坐标点")
        self.add_pump_curve_btn.setObjectName("ghostButton")
        self.remove_pump_curve_btn = QtWidgets.QPushButton("删除选中点")
        self.remove_pump_curve_btn.setObjectName("ghostButton")
        self.reset_pump_curve_btn = QtWidgets.QPushButton("重置水泵曲线")
        self.reset_pump_curve_btn.setObjectName("ghostButton")
        pump_actions.addWidget(self.add_pump_curve_btn)
        pump_actions.addWidget(self.remove_pump_curve_btn)
        pump_actions.addWidget(self.reset_pump_curve_btn)
        pump_actions.addStretch()
        pump_layout.addLayout(pump_actions)
        self.pump_curve_preview_label = QtWidgets.QLabel("当前温度：-- °C，预计水泵：--")
        self.pump_curve_preview_label.setObjectName("valuePill")
        pump_layout.addWidget(self.pump_curve_preview_label)
        pump_card.setMinimumHeight(450)
        top_row.addWidget(pump_card, 2)

        page_layout.addLayout(top_row)

        bottom_row = QtWidgets.QHBoxLayout()
        bottom_row.setSpacing(14)
        guide_card, guide_layout = self._create_panel("自动模式说明", "自动模式启用后，程序会按轮询到的温度同步更新风扇与水泵，并内置温度防抖。")
        guide_items = [
            "风扇曲线采用 20°C 到 100°C 区间，最高风扇百分比限制为 90%，最低支持 0%。",
            "水泵曲线支持 关闭 / 7V / 8V / 11V 四档，最大限制为 11V，运行时会自动贴合到最近档位。",
            "切回手动模式后，自动托管会停止；此时请按需要应用手动设置。",
            "自动模式支持温度回差、防抖采样次数，以及风扇/水泵独立最短启停间隔，减少频繁启停。",
            "最小化时会驻留托盘；点击关闭按钮将直接退出程序。",
        ]
        for item in guide_items:
            lbl = QtWidgets.QLabel(f"• {item}")
            lbl.setWordWrap(True)
            lbl.setObjectName("hintText")
            guide_layout.addWidget(lbl)
        guide_layout.addStretch()
        guide_card.setMinimumHeight(220)
        bottom_row.addWidget(guide_card, 3)

        status_card, status_layout = self._create_panel("自动模式状态", "显示当前自动模式的温度参考与预计输出。")
        self.auto_status_label = QtWidgets.QLabel("自动模式未启用")
        self.auto_status_label.setObjectName("hintText")
        self.auto_status_label.setWordWrap(True)
        status_layout.addWidget(self.auto_status_label)
        self.auto_runtime_label = QtWidgets.QLabel("当前温度：-- °C（控制值）")
        self.auto_runtime_label.setObjectName("valuePill")
        status_layout.addWidget(self.auto_runtime_label)

        auto_param_form = QtWidgets.QFormLayout()
        auto_param_form.setLabelAlignment(QtCore.Qt.AlignLeft)
        auto_param_form.setFormAlignment(QtCore.Qt.AlignTop)
        auto_param_form.setHorizontalSpacing(12)
        auto_param_form.setVerticalSpacing(10)

        self.auto_hysteresis_spin = QtWidgets.QSpinBox()
        self.auto_hysteresis_spin.setRange(1, 5)
        self.auto_hysteresis_spin.setSuffix(" °C")

        self.auto_samples_combo = QtWidgets.QComboBox()
        self.auto_samples_combo.addItem("3 次", 3)
        self.auto_samples_combo.addItem("5 次", 5)

        self.auto_fan_min_toggle_spin = QtWidgets.QDoubleSpinBox()
        self.auto_fan_min_toggle_spin.setRange(0.0, 30.0)
        self.auto_fan_min_toggle_spin.setDecimals(1)
        self.auto_fan_min_toggle_spin.setSingleStep(0.5)
        self.auto_fan_min_toggle_spin.setSuffix(" 秒")

        self.auto_pump_min_toggle_spin = QtWidgets.QDoubleSpinBox()
        self.auto_pump_min_toggle_spin.setRange(0.0, 30.0)
        self.auto_pump_min_toggle_spin.setDecimals(1)
        self.auto_pump_min_toggle_spin.setSingleStep(0.5)
        self.auto_pump_min_toggle_spin.setSuffix(" 秒")

        auto_param_form.addRow("温度回差", self.auto_hysteresis_spin)
        auto_param_form.addRow("防抖采样", self.auto_samples_combo)
        auto_param_form.addRow("风扇最短启停间隔", self.auto_fan_min_toggle_spin)
        auto_param_form.addRow("水泵最短启停间隔", self.auto_pump_min_toggle_spin)
        status_layout.addLayout(auto_param_form)

        self.auto_params_label = QtWidgets.QLabel("回差：2°C · 采样：3 次 · 风扇启停：3.0 秒 · 水泵启停：3.0 秒")
        self.auto_params_label.setObjectName("hintText")
        self.auto_params_label.setWordWrap(True)
        status_layout.addWidget(self.auto_params_label)
        status_layout.addStretch()
        status_card.setMinimumHeight(300)
        bottom_row.addWidget(status_card, 2)

        page_layout.addLayout(bottom_row)
        return auto

    def _update_connection_badge(self, connected: bool):
        self.connection_badge.setText("已连接" if connected else "未连接")
        self.connection_badge.setObjectName("statusBadgeConnected" if connected else "statusBadgeDisconnected")
        self.connection_badge.style().unpolish(self.connection_badge)
        self.connection_badge.style().polish(self.connection_badge)

    def _set_status_text(self, text, connected=None):
        self.status_label.setText(text)
        self.hero_status_label.setText(text)
        if connected is not None:
            self._update_connection_badge(connected)

    def _update_mode_hint(self):
        is_auto = self.mode_combo.currentIndex() == 1
        self.mode_badge.setText("自动模式" if is_auto else "手动模式")
        self.mode_hint_label.setText("按温度自动控制风扇与水泵" if is_auto else "直接设置风扇、水泵与灯效")
        self.preview_mode_label.setText("自动模式" if is_auto else "手动模式")

    def _fan_slider_text(self):
        value = int(self.fan_slider.value())
        if value <= 0:
            return "关闭"
        return f"{value}%"

    def _pump_slider_text(self):
        mapping = {
            0: "关闭",
            1: "7V",
            2: "8V",
            3: "11V",
        }
        return mapping.get(self.pump_slider.value(), "--")

    def _normalize_rgb_tuple(self, value, fallback=(255, 0, 0)):
        try:
            if isinstance(value, (list, tuple)) and len(value) == 3:
                return tuple(max(0, min(255, int(v))) for v in value)
        except Exception:
            pass
        return tuple(fallback)

    def _combo_color_value(self, combo, fallback=(255, 0, 0)):
        return self._normalize_rgb_tuple(combo.currentData(), fallback)

    def _set_combo_color_value(self, combo, color):
        target = self._normalize_rgb_tuple(color)
        for idx in range(combo.count()):
            try:
                item_value = self._normalize_rgb_tuple(combo.itemData(idx), target)
            except Exception:
                continue
            if item_value == target:
                combo.setCurrentIndex(idx)
                return True
        return False

    def _color_name_by_value(self, color):
        normalized = self._normalize_rgb_tuple(color)
        for name, value in COLOR_MAP.items():
            if self._normalize_rgb_tuple(value) == normalized:
                return name
        return f"RGB{normalized}"

    def _temperature_rgb_payload(self, temp=None):
        if temp is None:
            temp = self._current_control_temperature()
        low_limit = self.rgb_temp_low_spin.value() if hasattr(self, 'rgb_temp_low_spin') else self.settings.rgb_temp_threshold_low
        high_limit = self.rgb_temp_high_spin.value() if hasattr(self, 'rgb_temp_high_spin') else self.settings.rgb_temp_threshold_high
        if temp is None or temp <= low_limit:
            bucket = 'low'
            color = self._combo_color_value(self.rgb_temp_low_color, self.settings.rgb_temp_color_low) if hasattr(self, 'rgb_temp_low_color') else self.settings.rgb_temp_color_low
            bucket_text = '低温'
        elif temp <= high_limit:
            bucket = 'mid'
            color = self._combo_color_value(self.rgb_temp_mid_color, self.settings.rgb_temp_color_mid) if hasattr(self, 'rgb_temp_mid_color') else self.settings.rgb_temp_color_mid
            bucket_text = '中温'
        else:
            bucket = 'high'
            color = self._combo_color_value(self.rgb_temp_high_color, self.settings.rgb_temp_color_high) if hasattr(self, 'rgb_temp_high_color') else self.settings.rgb_temp_color_high
            bucket_text = '高温'
        mode = self.rgb_temp_mode_combo.currentData() if hasattr(self, 'rgb_temp_mode_combo') else self.settings.rgb_temp_mode
        return {
            'bucket': bucket,
            'bucket_text': bucket_text,
            'mode': mode,
            'mode_text': RGB_MODE_NAMES.get(mode.name, mode.name),
            'color': tuple(color),
            'color_text': self._color_name_by_value(color),
        }

    def _rgb_mode_text(self):
        if hasattr(self, 'rgb_temp_enabled_checkbox') and self.rgb_temp_enabled_checkbox.isChecked():
            payload = self._temperature_rgb_payload()
            return f"温控 · {payload['mode_text']} · {payload['bucket_text']} {payload['color_text']}"
        mode_text = self.rgb_mode.currentText() if hasattr(self, 'rgb_mode') else "--"
        if self.rgb_mode.currentData() in [RGBMode.OFF, RGBMode.COLORFUL, RGBMode.BREATHE_COLOR]:
            return mode_text
        return f"{mode_text} · {self.rgb_color.currentText()}"

    def _sync_rgb_input_states(self):
        temp_enabled = hasattr(self, 'rgb_temp_enabled_checkbox') and self.rgb_temp_enabled_checkbox.isChecked()
        manual_color_enabled = self.rgb_mode.currentData() not in [RGBMode.OFF, RGBMode.COLORFUL, RGBMode.BREATHE_COLOR]
        if hasattr(self, 'rgb_mode'):
            self.rgb_mode.setEnabled(not temp_enabled)
        if hasattr(self, 'rgb_color'):
            self.rgb_color.setEnabled((not temp_enabled) and manual_color_enabled)
        for widget in (
            getattr(self, 'rgb_temp_mode_combo', None),
            getattr(self, 'rgb_temp_low_spin', None),
            getattr(self, 'rgb_temp_high_spin', None),
            getattr(self, 'rgb_temp_low_color', None),
            getattr(self, 'rgb_temp_mid_color', None),
            getattr(self, 'rgb_temp_high_color', None),
        ):
            if widget is not None:
                widget.setEnabled(temp_enabled)

    def _current_control_temperature(self):
        temps = [t for t in (self.last_cpu_temp, self.last_gpu_temp) if t is not None]
        return max(temps) if temps else None

    def _update_control_temperature_history(self, temp):
        if temp is None:
            return
        self._control_temp_history.append(float(temp))

    def _auto_control_temperature(self):
        if self._control_temp_history:
            try:
                return float(median(self._control_temp_history))
            except Exception:
                pass
        return self._current_control_temperature()

    def _first_nonzero_curve_temp(self, widget):
        if widget is None:
            return None
        for temp in range(20, 101):
            try:
                if int(widget.interpolate(temp)) > 0:
                    return temp
            except Exception:
                break
        return None

    def _stabilize_auto_targets(self, control_temp, fan_percent, pump_value):
        if control_temp is None:
            return fan_percent, pump_value

        fan_percent = int(fan_percent)
        pump_value = int(pump_value)
        hysteresis_c = float(getattr(self.settings, 'auto_hysteresis_c', DEFAULT_AUTO_HYSTERESIS_C))
        fan_min_toggle_interval = float(getattr(self.settings, 'auto_fan_min_toggle_interval_sec', DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC))
        pump_min_toggle_interval = float(getattr(self.settings, 'auto_pump_min_toggle_interval_sec', DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC))

        fan_activation_temp = self._first_nonzero_curve_temp(getattr(self, 'curve_widget', None))
        if self._auto_applied_fan_percent and self._auto_applied_fan_percent > 0 and fan_percent <= 0 and fan_activation_temp is not None:
            if control_temp > fan_activation_temp - hysteresis_c:
                fan_percent = max(1, int(self.curve_widget.interpolate(fan_activation_temp)))

        pump_activation_temp = self._first_nonzero_curve_temp(getattr(self, 'pump_curve_widget', None))
        if self._auto_applied_pump_value and self._auto_applied_pump_value > 0 and pump_value <= 0 and pump_activation_temp is not None:
            if control_temp > pump_activation_temp - hysteresis_c:
                hold_value = int(self.pump_curve_widget.interpolate(pump_activation_temp))
                pump_value = hold_value if hold_value > 0 else 7

        now = time.monotonic()

        previous_fan_on = bool(self._auto_applied_fan_percent and self._auto_applied_fan_percent > 0)
        target_fan_on = fan_percent > 0
        if self._auto_applied_fan_percent is not None and previous_fan_on != target_fan_on:
            if self._last_fan_toggle_ts and (now - self._last_fan_toggle_ts) < fan_min_toggle_interval:
                fan_percent = int(self._auto_applied_fan_percent)
            else:
                self._last_fan_toggle_ts = now

        previous_pump_on = bool(self._auto_applied_pump_value and self._auto_applied_pump_value > 0)
        target_pump_on = pump_value > 0
        if self._auto_applied_pump_value is not None and previous_pump_on != target_pump_on:
            if self._last_pump_toggle_ts and (now - self._last_pump_toggle_ts) < pump_min_toggle_interval:
                pump_value = int(self._auto_applied_pump_value)
            else:
                self._last_pump_toggle_ts = now

        return fan_percent, pump_value

    def _auto_fan_percent(self):
        temp = self._auto_control_temperature()
        if temp is None or not hasattr(self, 'curve_widget'):
            return None
        raw_fan = int(self.curve_widget.interpolate(temp))
        raw_pump = int(self.pump_curve_widget.interpolate(temp)) if hasattr(self, 'pump_curve_widget') else 0
        fan_percent, _ = self._stabilize_auto_targets(temp, raw_fan, raw_pump)
        return fan_percent

    def _auto_pump_voltage_text(self):
        temp = self._auto_control_temperature()
        if temp is None or not hasattr(self, 'pump_curve_widget'):
            return None
        raw_fan = int(self.curve_widget.interpolate(temp)) if hasattr(self, 'curve_widget') else 0
        raw_pump = int(self.pump_curve_widget.interpolate(temp))
        _, pump_value = self._stabilize_auto_targets(temp, raw_fan, raw_pump)
        return pump_curve_value_to_text(pump_value)

    def _save_curve_settings_from_ui(self):
        self.settings.fan_curve_points = normalize_fan_curve_points(self.curve_widget.points)
        self.settings.pump_curve_points = normalize_pump_curve_points(self.pump_curve_widget.points)
        self.settings.selected_mode_index = self.mode_combo.currentIndex()

    def _save_rgb_temp_settings_from_ui(self):
        self.settings.rgb_temp_enabled = self.rgb_temp_enabled_checkbox.isChecked()
        self.settings.rgb_temp_mode = self.rgb_temp_mode_combo.currentData()
        low = self.rgb_temp_low_spin.value()
        high = max(low + 1, self.rgb_temp_high_spin.value())
        self.rgb_temp_high_spin.blockSignals(True)
        self.rgb_temp_high_spin.setValue(high)
        self.rgb_temp_high_spin.blockSignals(False)
        self.settings.rgb_temp_threshold_low = low
        self.settings.rgb_temp_threshold_high = high
        self.settings.rgb_temp_color_low = self._combo_color_value(self.rgb_temp_low_color, self.settings.rgb_temp_color_low)
        self.settings.rgb_temp_color_mid = self._combo_color_value(self.rgb_temp_mid_color, self.settings.rgb_temp_color_mid)
        self.settings.rgb_temp_color_high = self._combo_color_value(self.rgb_temp_high_color, self.settings.rgb_temp_color_high)

    def _refresh_control_temp_history_maxlen(self):
        maxlen = 5 if int(getattr(self.settings, 'auto_debounce_samples', DEFAULT_AUTO_DEBOUNCE_SAMPLES)) >= 5 else 3
        old_values = list(self._control_temp_history) if hasattr(self, '_control_temp_history') else []
        self._control_temp_history = deque(old_values[-maxlen:], maxlen=maxlen)

    def _save_auto_debounce_settings_from_ui(self):
        if not hasattr(self, 'auto_hysteresis_spin'):
            return
        self.settings.auto_hysteresis_c = int(self.auto_hysteresis_spin.value())
        self.settings.auto_debounce_samples = int(self.auto_samples_combo.currentData() or DEFAULT_AUTO_DEBOUNCE_SAMPLES)
        self.settings.auto_fan_min_toggle_interval_sec = float(self.auto_fan_min_toggle_spin.value())
        self.settings.auto_pump_min_toggle_interval_sec = float(self.auto_pump_min_toggle_spin.value())
        self._refresh_control_temp_history_maxlen()

    def on_auto_debounce_settings_changed(self, *args):
        if getattr(self, '_syncing_ui', False):
            return
        self._save_auto_debounce_settings_from_ui()
        self._update_control_summaries()
        self.settings.save()

    def _update_control_summaries(self):
        if not hasattr(self, 'fan_slider'):
            return
        fan_text = self._fan_slider_text()
        pump_text = self._pump_slider_text()
        rgb_text = self._rgb_mode_text() if hasattr(self, 'rgb_mode') else "--"
        self.fan_value_pill.setText(fan_text)
        self.pump_value_pill.setText(pump_text)
        self.rgb_value_pill.setText(rgb_text)

        is_auto = hasattr(self, 'mode_combo') and self.mode_combo.currentIndex() == 1
        auto_temp = self._current_control_temperature()
        auto_fan_pct = self._auto_fan_percent()
        auto_pump_text = self._auto_pump_voltage_text()

        if is_auto and auto_fan_pct is not None:
            self.fan_summary_label.setText(f"自动 {'关闭' if auto_fan_pct <= 0 else f'{auto_fan_pct}%'}")
            self.preview_fan_label.setText(f"自动 / {'关闭' if auto_fan_pct <= 0 else f'{auto_fan_pct}%'}")
        else:
            self.fan_summary_label.setText(fan_text)
            self.preview_fan_label.setText(f"{fan_text} / 固定")

        if is_auto and auto_pump_text is not None:
            self.pump_summary_label.setText(f"自动 {auto_pump_text}")
            self.preview_pump_label.setText(f"自动 / {auto_pump_text}")
        else:
            self.pump_summary_label.setText(pump_text)
            self.preview_pump_label.setText(f"{pump_text} / 固定")

        self.preview_rgb_label.setText(rgb_text.replace(' · ', ' / '))

        if hasattr(self, 'fan_curve_preview_label'):
            if auto_temp is not None and auto_fan_pct is not None:
                self.fan_curve_preview_label.setText(f"当前温度：{int(auto_temp)} °C，预计风扇：{'关闭' if auto_fan_pct <= 0 else f'{auto_fan_pct} %'}")
            else:
                self.fan_curve_preview_label.setText("当前温度：-- °C，预计风扇：-- %")
        if hasattr(self, 'pump_curve_preview_label'):
            if auto_temp is not None and auto_pump_text is not None:
                self.pump_curve_preview_label.setText(f"当前温度：{int(auto_temp)} °C，预计水泵：{auto_pump_text}")
            else:
                self.pump_curve_preview_label.setText("当前温度：-- °C，预计水泵：--")
        if hasattr(self, 'auto_runtime_label'):
            if auto_temp is not None:
                self.auto_runtime_label.setText(f"当前温度：{int(auto_temp)} °C（控制值）")
            else:
                self.auto_runtime_label.setText("当前温度：-- °C（控制值）")
        if hasattr(self, 'auto_params_label'):
            self.auto_params_label.setText(
                f"回差：{int(getattr(self.settings, 'auto_hysteresis_c', DEFAULT_AUTO_HYSTERESIS_C))}°C · "
                f"采样：{int(getattr(self.settings, 'auto_debounce_samples', DEFAULT_AUTO_DEBOUNCE_SAMPLES))} 次 · "
                f"风扇启停：{float(getattr(self.settings, 'auto_fan_min_toggle_interval_sec', DEFAULT_AUTO_FAN_MIN_TOGGLE_INTERVAL_SEC)):.1f} 秒 · "
                f"水泵启停：{float(getattr(self.settings, 'auto_pump_min_toggle_interval_sec', DEFAULT_AUTO_PUMP_MIN_TOGGLE_INTERVAL_SEC)):.1f} 秒"
            )

    def _reset_auto_control_state(self):
        self._control_temp_history.clear()
        self._auto_applied_fan_percent = None
        self._auto_applied_pump_value = None
        self._last_fan_toggle_ts = 0.0
        self._last_pump_toggle_ts = 0.0

    def on_rgb_mode_changed(self, index):
        if getattr(self, '_syncing_ui', False):
            return
        self._sync_rgb_input_states()
        self._update_control_summaries()

    def on_rgb_temp_controls_changed(self, *args):
        if getattr(self, '_syncing_ui', False):
            return
        if hasattr(self, 'rgb_temp_low_spin') and hasattr(self, 'rgb_temp_high_spin'):
            if self.rgb_temp_high_spin.value() <= self.rgb_temp_low_spin.value():
                self.rgb_temp_high_spin.blockSignals(True)
                self.rgb_temp_high_spin.setValue(self.rgb_temp_low_spin.value() + 1)
                self.rgb_temp_high_spin.blockSignals(False)
        self._save_rgb_temp_settings_from_ui()
        self._sync_rgb_input_states()
        self._update_control_summaries()
        self.settings.save()

    def on_device_selection_changed(self, index):
        if not hasattr(self, 'device_combo'):
            return
        device_text = self.device_combo.currentText().strip()
        if device_text:
            self.selected_device_label.setText(f"设备：{device_text}")
            self.device_tip_label.setText("已选择设备，点击连接即可建立蓝牙连接。")
            if not (self.client and self.client.is_connected):
                self._set_status_text(UI['select_prompt'], connected=False)
                self._refresh_export_api_state()
            self._update_connection_controls()
        else:
            self.selected_device_label.setText("设备：等待扫描")
            self.device_tip_label.setText("当前未发现可连接设备")
            if not (self.client and self.client.is_connected):
                self._set_status_text(UI['searching'], connected=False)
                self._refresh_export_api_state()

    def _curve_insert_position(self, points, selected_index=None):
        if len(points) < 2:
            return None
        if selected_index is not None:
            if 0 <= selected_index < len(points) - 1:
                return selected_index
            if 1 <= selected_index < len(points):
                return selected_index - 1
        largest_gap = None
        insert_after = None
        for idx in range(len(points) - 1):
            gap = int(points[idx + 1][0]) - int(points[idx][0])
            if gap <= 1:
                continue
            if largest_gap is None or gap > largest_gap:
                largest_gap = gap
                insert_after = idx
        return insert_after

    def _insert_curve_point(self, widget, normalizer):
        points = [tuple(point) for point in widget.points]
        if len(points) >= MAX_CURVE_POINTS:
            return False
        insert_after = self._curve_insert_position(points, widget.selected)
        if insert_after is None:
            return False
        t0, v0 = points[insert_after]
        t1, v1 = points[insert_after + 1]
        new_temp = (int(t0) + int(t1)) // 2
        if new_temp <= int(t0):
            new_temp = int(t0) + 1
        if new_temp >= int(t1):
            new_temp = int(t1) - 1
        if new_temp <= int(t0) or new_temp >= int(t1):
            return False
        new_value = normalizer(round((int(v0) + int(v1)) / 2))
        points.insert(insert_after + 1, (new_temp, new_value))
        widget.points = points
        widget.selected = insert_after + 1
        widget.update()
        self._on_curve_points_edited()
        return True

    def _remove_curve_point(self, widget):
        points = [tuple(point) for point in widget.points]
        if len(points) <= MIN_CURVE_POINTS or widget.selected is None:
            return False
        del points[widget.selected]
        widget.points = points
        widget.selected = min(widget.selected, len(points) - 1) if points else None
        widget.update()
        self._on_curve_points_edited()
        return True

    def add_fan_curve_point(self, checked=False):
        self._insert_curve_point(self.curve_widget, clamp_curve_percent)

    def remove_fan_curve_point(self, checked=False):
        self._remove_curve_point(self.curve_widget)

    def add_pump_curve_point(self, checked=False):
        self._insert_curve_point(self.pump_curve_widget, clamp_pump_curve_value)

    def remove_pump_curve_point(self, checked=False):
        self._remove_curve_point(self.pump_curve_widget)

    def _update_curve_editor_buttons(self):
        if hasattr(self, 'add_fan_curve_btn'):
            self.add_fan_curve_btn.setEnabled(len(self.curve_widget.points) < MAX_CURVE_POINTS)
        if hasattr(self, 'remove_fan_curve_btn'):
            self.remove_fan_curve_btn.setEnabled(
                self.curve_widget.selected is not None and len(self.curve_widget.points) > MIN_CURVE_POINTS
            )
        if hasattr(self, 'add_pump_curve_btn'):
            self.add_pump_curve_btn.setEnabled(len(self.pump_curve_widget.points) < MAX_CURVE_POINTS)
        if hasattr(self, 'remove_pump_curve_btn'):
            self.remove_pump_curve_btn.setEnabled(
                self.pump_curve_widget.selected is not None and len(self.pump_curve_widget.points) > MIN_CURVE_POINTS
            )

    def _on_curve_points_edited(self):
        self.apply_curve_btn.setEnabled(True)
        self._save_curve_settings_from_ui()
        self._update_control_summaries()
        self._update_curve_editor_buttons()
        self.settings.save()

    def reset_curve_points(self, checked=False):
        self.reset_fan_curve_points()

    def reset_fan_curve_points(self, checked=False):
        self.fan_curve_points = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
        if hasattr(self, 'curve_widget'):
            self.curve_widget.points = [tuple(point) for point in DEFAULT_FAN_CURVE_POINTS]
            self.curve_widget.selected = None
            self.curve_widget.update()
        self._save_curve_settings_from_ui()
        self._update_control_summaries()
        self._update_curve_editor_buttons()
        self.settings.save()

    def reset_pump_curve_points(self, checked=False):
        self.pump_curve_points = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
        if hasattr(self, 'pump_curve_widget'):
            self.pump_curve_widget.points = [tuple(point) for point in DEFAULT_PUMP_CURVE_POINTS]
            self.pump_curve_widget.selected = None
            self.pump_curve_widget.update()
        self._save_curve_settings_from_ui()
        self._update_control_summaries()
        self._update_curve_editor_buttons()
        self.settings.save()

    def _build_ui(self):
        self._apply_styles()
        self.resize(1060, 820)
        self.setMinimumSize(920, 700)
        self.setWindowState(self.windowState() | QtCore.Qt.WindowMaximized)

        main = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(main)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(16)

        header_card = QtWidgets.QFrame()
        header_card.setObjectName("heroCard")
        header_layout = QtWidgets.QHBoxLayout(header_card)
        header_layout.setContentsMargins(24, 20, 24, 20)
        header_layout.setSpacing(18)

        header_text = QtWidgets.QVBoxLayout()
        header_text.setSpacing(4)
        self.window_title_label = QtWidgets.QLabel("水冷管理器")
        self.window_title_label.setObjectName("heroTitle")
        self.window_subtitle_label = QtWidgets.QLabel("更清晰的连接状态、散热控制和灯效管理")
        self.window_subtitle_label.setObjectName("heroSubtitle")
        header_text.addWidget(self.window_title_label)
        header_text.addWidget(self.window_subtitle_label)
        header_text.addStretch()
        header_layout.addLayout(header_text, 1)

        hero_right = QtWidgets.QVBoxLayout()
        hero_right.setSpacing(10)
        badge_row = QtWidgets.QHBoxLayout()
        badge_row.setSpacing(10)
        self.connection_badge = QtWidgets.QLabel("未连接")
        self.connection_badge.setObjectName("statusBadgeDisconnected")
        self.mode_badge = QtWidgets.QLabel("手动模式")
        self.mode_badge.setObjectName("secondaryBadge")
        badge_row.addStretch()
        badge_row.addWidget(self.mode_badge)
        badge_row.addWidget(self.connection_badge)
        hero_right.addLayout(badge_row)
        self.selected_device_label = QtWidgets.QLabel("设备：等待扫描")
        self.selected_device_label.setObjectName("heroMeta")
        self.hero_status_label = QtWidgets.QLabel(UI['searching'])
        self.hero_status_label.setObjectName("heroMeta")
        self.hero_status_label.setWordWrap(True)
        hero_right.addWidget(self.selected_device_label, 0, QtCore.Qt.AlignRight)
        hero_right.addWidget(self.hero_status_label, 0, QtCore.Qt.AlignRight)
        header_layout.addLayout(hero_right)
        layout.addWidget(header_card)

        stats_row = QtWidgets.QHBoxLayout()
        stats_row.setSpacing(12)
        self.cpu_value_label = self._create_info_card(stats_row, "CPU 温度", "-- °C")
        self.gpu_value_label = self._create_info_card(stats_row, "显卡温度", "-- °C")
        self.fan_summary_label = self._create_info_card(stats_row, "风扇功率", "30%")
        self.pump_summary_label = self._create_info_card(stats_row, "水泵状态", "7V")
        layout.addLayout(stats_row)

        top_bar = QtWidgets.QHBoxLayout()
        top_bar.setSpacing(14)
        mode_wrap = QtWidgets.QFrame()
        mode_wrap.setObjectName("toolbarCard")
        mode_layout = QtWidgets.QHBoxLayout(mode_wrap)
        mode_layout.setContentsMargins(18, 12, 18, 12)
        mode_layout.setSpacing(12)
        mode_caption = QtWidgets.QLabel("工作模式")
        mode_caption.setObjectName("sectionMiniTitle")
        self.mode_combo = QtWidgets.QComboBox()
        self.mode_combo.addItem(UI['mode_manual'])
        self.mode_combo.addItem(UI['mode_curve'])
        self.mode_hint_label = QtWidgets.QLabel("直接设置风扇、水泵与灯效")
        self.mode_hint_label.setObjectName("mutedText")
        mode_layout.addWidget(mode_caption)
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addWidget(self.mode_hint_label, 1)
        top_bar.addWidget(mode_wrap, 1)
        layout.addLayout(top_bar)

        self.pages = QtWidgets.QStackedWidget()
        self.pages.setObjectName("pages")
        self.pages.addWidget(self._wrap_page_in_scroll(self._build_manual_page()))
        self.pages.addWidget(self._wrap_page_in_scroll(self._build_auto_page()))
        layout.addWidget(self.pages, 1)

        footer_top = QtWidgets.QHBoxLayout()
        footer_top.setSpacing(12)
        status_card = QtWidgets.QFrame()
        status_card.setObjectName("footerCard")
        status_layout = QtWidgets.QHBoxLayout(status_card)
        status_layout.setContentsMargins(16, 12, 16, 12)
        status_layout.setSpacing(12)
        status_tag = QtWidgets.QLabel("状态")
        status_tag.setObjectName("statusTag")
        self.status_label = QtWidgets.QLabel(UI['searching'])
        self.status_label.setWordWrap(True)
        self.status_label.setObjectName("statusText")
        status_layout.addWidget(status_tag, 0, QtCore.Qt.AlignTop)
        status_layout.addWidget(self.status_label, 1)
        status_layout.addStretch()
        self.disconnect_btn = QtWidgets.QPushButton(UI['btn_disconnect'])
        self.disconnect_btn.setObjectName("ghostButton")
        self.disconnect_btn.setEnabled(False)
        status_layout.addWidget(self.disconnect_btn, 0, QtCore.Qt.AlignVCenter)
        footer_top.addWidget(status_card, 1)

        metrics_card = QtWidgets.QFrame()
        metrics_card.setObjectName("footerCard")
        metrics_layout = QtWidgets.QHBoxLayout(metrics_card)
        metrics_layout.setContentsMargins(16, 12, 16, 12)
        metrics_layout.setSpacing(12)
        metrics_layout.addWidget(QtWidgets.QLabel(UI['label_update_speed']))
        self.update_speed_combo = QtWidgets.QComboBox()
        for sec, label in self.UPDATE_INTERVALS:
            self.update_speed_combo.addItem(label, sec)
        self.update_speed_combo.addItem("自定义", 'custom')
        self.update_speed_custom = QtWidgets.QDoubleSpinBox()
        self.update_speed_custom.setRange(0.2, 60.0)
        self.update_speed_custom.setDecimals(1)
        self.update_speed_custom.setSingleStep(0.5)
        self.update_speed_custom.setSuffix(" 秒")
        self.update_speed_custom.setFixedWidth(110)
        metrics_layout.addWidget(self.update_speed_combo)
        metrics_layout.addWidget(self.update_speed_custom)
        self.update_speed_custom.hide()
        metrics_layout.addStretch()
        footer_top.addWidget(metrics_card)
        layout.addLayout(footer_top)

        footer_bottom = QtWidgets.QHBoxLayout()
        footer_bottom.setSpacing(12)
        options_card = QtWidgets.QFrame()
        options_card.setObjectName("footerCard")
        options_layout = QtWidgets.QHBoxLayout(options_card)
        options_layout.setContentsMargins(16, 12, 16, 12)
        options_layout.setSpacing(18)
        self.auto_connect_checkbox = QtWidgets.QCheckBox(UI['label_auto_connect'])
        self.auto_connect_checkbox.setChecked(self.settings.auto_connect)
        self.auto_connect_checkbox.stateChanged.connect(self.on_auto_connect_changed)
        self.auto_start_checkbox = QtWidgets.QCheckBox(UI['label_auto_start'])
        self.auto_start_checkbox.setChecked(self.settings.auto_start)
        self.auto_start_checkbox.stateChanged.connect(self.on_auto_start_changed)
        options_layout.addWidget(self.auto_connect_checkbox)
        options_layout.addWidget(self.auto_start_checkbox)
        self.export_api_enable_checkbox = QtWidgets.QCheckBox("启用 API")
        options_layout.addWidget(self.export_api_enable_checkbox)
        self.export_api_port_label = QtWidgets.QLabel("API 端口")
        options_layout.addWidget(self.export_api_port_label)
        self.export_api_port_spin = QtWidgets.QSpinBox()
        self.export_api_port_spin.setRange(1024, 65535)
        self.export_api_port_spin.setFixedWidth(96)
        options_layout.addWidget(self.export_api_port_spin)
        self.export_api_status_label = QtWidgets.QLabel("")
        self.export_api_status_label.setObjectName("mutedText")
        options_layout.addWidget(self.export_api_status_label)
        self.dingtalk_enable_checkbox = QtWidgets.QCheckBox("钉钉推送")
        options_layout.addWidget(self.dingtalk_enable_checkbox)
        self.dingtalk_settings_btn = QtWidgets.QPushButton("设置")
        self.dingtalk_settings_btn.setFixedHeight(34)
        self.dingtalk_settings_btn.setFixedWidth(76)
        options_layout.addWidget(self.dingtalk_settings_btn)
        options_layout.addWidget(QtWidgets.QLabel("主题"))
        self.theme_combo = QtWidgets.QComboBox()
        self.theme_combo.addItem("深色", 'dark')
        self.theme_combo.addItem("浅色", 'light')
        self.theme_combo.addItem("跟随系统", 'system')
        options_layout.addWidget(self.theme_combo)
        options_layout.addStretch()
        footer_bottom.addWidget(options_card)
        layout.addLayout(footer_bottom)

        self.setCentralWidget(main)
        self.mode_combo.currentIndexChanged.connect(self.on_mode_changed)
        self.connect_btn.clicked.connect(self.connect_device)
        self.rescan_btn.clicked.connect(self.trigger_rescan)
        self.disconnect_btn.clicked.connect(self.disconnect_device)
        self.device_combo.currentIndexChanged.connect(self.on_device_selection_changed)
        self.apply_manual_btn.clicked.connect(self.apply_fan_and_pump)
        self.apply_rgb_btn.clicked.connect(self.apply_rgb)
        self.apply_all_btn.clicked.connect(self.apply_all)
        self.apply_curve_btn.clicked.connect(self.apply_curve)
        self.preset_silent_btn.clicked.connect(self.apply_silent_preset)
        self.preset_balanced_btn.clicked.connect(self.apply_balanced_preset)
        self.preset_performance_btn.clicked.connect(self.apply_performance_preset)
        self.add_fan_curve_btn.clicked.connect(self.add_fan_curve_point)
        self.remove_fan_curve_btn.clicked.connect(self.remove_fan_curve_point)
        self.reset_curve_btn.clicked.connect(self.reset_fan_curve_points)
        self.add_pump_curve_btn.clicked.connect(self.add_pump_curve_point)
        self.remove_pump_curve_btn.clicked.connect(self.remove_pump_curve_point)
        self.reset_pump_curve_btn.clicked.connect(self.reset_pump_curve_points)
        self.rgb_mode.currentIndexChanged.connect(self.on_rgb_mode_changed)
        self.rgb_color.currentIndexChanged.connect(lambda _: self._update_control_summaries())
        self.rgb_temp_enabled_checkbox.stateChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_mode_combo.currentIndexChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_low_spin.valueChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_high_spin.valueChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_low_color.currentIndexChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_mid_color.currentIndexChanged.connect(self.on_rgb_temp_controls_changed)
        self.rgb_temp_high_color.currentIndexChanged.connect(self.on_rgb_temp_controls_changed)
        self.fan_slider.valueChanged.connect(self.on_manual_control_value_changed)
        self.pump_slider.valueChanged.connect(self.on_manual_control_value_changed)
        self.update_speed_combo.currentIndexChanged.connect(self.update_interval_changed)
        self.update_speed_custom.valueChanged.connect(self.update_interval_changed)
        self.auto_hysteresis_spin.valueChanged.connect(self.on_auto_debounce_settings_changed)
        self.auto_samples_combo.currentIndexChanged.connect(self.on_auto_debounce_settings_changed)
        self.auto_fan_min_toggle_spin.valueChanged.connect(self.on_auto_debounce_settings_changed)
        self.auto_pump_min_toggle_spin.valueChanged.connect(self.on_auto_debounce_settings_changed)
        self.export_api_enable_checkbox.stateChanged.connect(self.on_export_api_settings_changed)
        self.export_api_port_spin.valueChanged.connect(self.on_export_api_settings_changed)
        self.dingtalk_enable_checkbox.stateChanged.connect(self.on_dingtalk_enable_changed)
        self.dingtalk_settings_btn.clicked.connect(self.open_dingtalk_settings_dialog)
        self.theme_combo.currentIndexChanged.connect(self.on_theme_changed)
        self.pages.setCurrentIndex(0)
        self._apply_theme()
        self._update_dingtalk_controls()
        self._update_mode_hint()
        self._sync_rgb_input_states()
        self._update_control_summaries()
        self._update_curve_editor_buttons()


    def _set_manual_action_buttons_enabled(self, connected: bool):
        for btn in (
            self.apply_manual_btn,
            self.apply_rgb_btn,
            self.apply_all_btn,
            self.apply_curve_btn,
            self.preset_silent_btn,
            self.preset_balanced_btn,
            self.preset_performance_btn,
        ):
            btn.setEnabled(bool(connected))
        self._update_manual_apply_hint()

    def _current_manual_control_signature(self):
        if not hasattr(self, 'fan_slider') or not hasattr(self, 'pump_slider'):
            return None
        return int(self.fan_slider.value()), int(self.pump_slider.value())

    def _saved_manual_control_signature(self):
        fan_value = 0 if self.settings.fan_is_off else self._duty_to_fan_slider(self.settings.current_fan_speed)
        pump_value = 0 if self.settings.pump_is_off else self._voltage_to_pump_slider(self.settings.current_voltage)
        return int(fan_value), int(pump_value)

    def _manual_controls_dirty(self):
        current = self._current_manual_control_signature()
        if current is None:
            return False
        return current != self._saved_manual_control_signature()

    def _update_manual_apply_hint(self):
        if not hasattr(self, 'manual_apply_hint_label') or not hasattr(self, 'apply_manual_btn'):
            return
        dirty = (
            not getattr(self, '_syncing_ui', False)
            and not getattr(self, '_manual_prompt_suspend', False)
            and self._manual_controls_dirty()
        )
        connected = bool(self.client and self.client.is_connected)
        manual_mode = not hasattr(self, 'mode_combo') or self.mode_combo.currentIndex() == 0
        if dirty and manual_mode:
            self.manual_apply_hint_label.setText("风扇/水泵参数已变更，点击“应用”按钮后才会生效。")
            self.apply_manual_btn.setText("应用（待生效）")
        else:
            if connected and manual_mode:
                self.manual_apply_hint_label.setText("拖动风扇或水泵后，需要点击“应用”按钮才会生效。")
            elif manual_mode:
                self.manual_apply_hint_label.setText("连接设备后，可使用快捷模式或手动调节风扇/水泵并点击“应用”。")
            else:
                self.manual_apply_hint_label.setText("当前处于自动模式；切回手动模式后，可使用快捷模式或手动应用风扇/水泵设置。")
            self.apply_manual_btn.setText(UI['btn_apply_manual'])

    def on_manual_control_value_changed(self, *args):
        self._update_control_summaries()
        self._update_manual_apply_hint()

    async def _apply_manual_preset_values(self, fan_percent: int, pump_slider_value: int):
        self._manual_prompt_suspend = True
        try:
            self.fan_slider.setValue(int(fan_percent))
            self.pump_slider.setValue(int(pump_slider_value))
            self._update_control_summaries()
        finally:
            self._manual_prompt_suspend = False
        await self.apply_fan_and_pump()
        self._update_manual_apply_hint()

    @asyncSlot()
    async def apply_silent_preset(self):
        await self._apply_manual_preset_values(30, 1)

    @asyncSlot()
    async def apply_balanced_preset(self):
        await self._apply_manual_preset_values(60, 2)

    @asyncSlot()
    async def apply_performance_preset(self):
        await self._apply_manual_preset_values(90, 3)

    def _last_known_device_address(self):
        value = getattr(self.settings, 'last_device_address', None)
        value = str(value).strip() if value else ''
        return value or None

    def _last_known_device_name(self):
        value = getattr(self.settings, 'last_device_name', None)
        value = str(value).strip() if value else ''
        return value or '上次连接设备'

    def _queue_disconnect_handler(self, reason=None):
        if getattr(self, '_disconnect_callback_scheduled', False) or self._is_exiting:
            logging.info('BLE disconnect callback ignored: already scheduled=%s, exiting=%s', getattr(self, '_disconnect_callback_scheduled', False), self._is_exiting)
            return
        self._disconnect_callback_scheduled = True
        logging.info('BLE disconnect callback queued: reason=%s', reason or 'unknown')

        async def _run():
            try:
                await self._handle_unexpected_disconnect(reason=reason)
            finally:
                self._disconnect_callback_scheduled = False

        QtCore.QTimer.singleShot(0, lambda: asyncio.ensure_future(_run()))

    def _update_connection_controls(self):
        has_device = bool(hasattr(self, 'device_combo') and self.device_combo.count() > 0 and self.device_combo.currentData())
        connected = bool(self.client and self.client.is_connected)
        busy = bool(self.is_connecting or self.is_disconnecting or self._is_exiting)
        scanning = bool(getattr(self, 'is_scanning', False))
        if hasattr(self, 'connect_btn'):
            self.connect_btn.setEnabled((not connected) and has_device and (not busy) and (not scanning))
        if hasattr(self, 'rescan_btn'):
            self.rescan_btn.setEnabled((not busy) and (not scanning))
            self.rescan_btn.setText("扫描中..." if scanning else "重新扫描")
        if hasattr(self, 'disconnect_btn'):
            self.disconnect_btn.setEnabled(connected and (not busy))
        if not connected:
            self._set_manual_action_buttons_enabled(False)

    def _persist_ui_settings(self):
        self.settings.selected_mode_index = self.mode_combo.currentIndex() if hasattr(self, 'mode_combo') else 0
        self.settings.auto_mode_enabled = bool(self.auto_mode_active)
        if hasattr(self, 'curve_widget') and hasattr(self, 'pump_curve_widget'):
            self._save_curve_settings_from_ui()
        if hasattr(self, 'rgb_temp_enabled_checkbox'):
            self._save_rgb_temp_settings_from_ui()
        if hasattr(self, 'export_api_enable_checkbox'):
            self.settings.export_api_enabled = self.export_api_enable_checkbox.isChecked()
            self.settings.export_api_port = int(self.export_api_port_spin.value())
        if hasattr(self, 'dingtalk_enable_checkbox'):
            self.settings.dingtalk_webhook_enabled = self.dingtalk_enable_checkbox.isChecked()
        self.settings.save()

    async def _disconnect_client(self, send_reset: bool = True, update_status: bool = True):
        client = self.client
        self.client = None
        try:
            if client and client.is_connected:
                if send_reset:
                    try:
                        await write_reset(client)
                        await asyncio.sleep(0.15)
                    except Exception:
                        pass
                try:
                    await client.disconnect()
                except Exception:
                    pass
        finally:
            self.pump_runtime_on = False
            self.pump_runtime_voltage = None
            self.auto_mode_active = False
            self._last_temp_rgb_bucket = None
            self.export_api_state.update(connected=False, device_name=None)
            self._refresh_export_api_state()
            self._update_connection_controls()
            if update_status and not self._is_exiting:
                self.device_tip_label.setText("设备已断开，可重新连接。")
                logging.info('BLE disconnected')
                self._set_status_text("设备已断开连接", connected=False)
                self.auto_status_label.setText('自动模式未启用')

    async def _shutdown_and_quit(self):
        if self._is_exiting:
            return
        self._is_exiting = True
        logging.info('Application shutdown requested')
        if hasattr(self, 'temp_timer') and self.temp_timer.isActive():
            self.temp_timer.stop()
            logging.info('Temperature timer stopped for shutdown')
        self._persist_ui_settings()
        self.device_tip_label.setText("正在断开设备并退出…")
        self._set_status_text("正在断开设备并退出...", connected=bool(self.client and self.client.is_connected))
        self._update_connection_controls()
        await self._disconnect_client(send_reset=True, update_status=False)
        await asyncio.sleep(0.05)
        if self.export_api_server:
            self.export_api_server.stop()
        if self.tray_icon:
            self.tray_icon.hide()
        QtWidgets.qApp.quit()

    async def _handle_unexpected_disconnect(self, reason=None):
        if self._is_exiting or self.is_disconnecting:
            return
        logging.warning('BLE device disconnected unexpectedly: %s', reason or 'unknown')
        await self._disconnect_client(send_reset=False, update_status=True)
        self._auto_reconnect_pending = True
        self._notify_connection_event('unexpected_disconnect')
        last_addr = self._last_known_device_address()
        last_name = self._last_known_device_name()
        if last_addr and self.device_combo.count() == 0:
            self.device_combo.addItem(f"{last_name} [{last_addr}]（上次连接）", last_addr)
            self.device_combo.setEnabled(True)
        self.device_tip_label.setText("蓝牙连接意外断开，程序已停止自动控制；可直接重连上次设备。")
        self._set_status_text("蓝牙连接意外断开", connected=False)
        self._update_connection_controls()
        if not self.is_scanning:
            logging.info('Schedule BLE rescan after unexpected disconnect')
            QtCore.QTimer.singleShot(3000, lambda: asyncio.ensure_future(self.scan_and_populate()))

    def _is_auto_mode_selected(self):
        return bool(hasattr(self, 'mode_combo') and self.mode_combo.currentIndex() == 1)

    def _current_export_fan_percent(self):
        if self.auto_mode_active and self._auto_applied_fan_percent is not None:
            return max(0, int(self._auto_applied_fan_percent))
        # 自动模式已选中但尚未正式启用时，API 输出当前温度下的预估值，便于外部面板展示。
        if self._is_auto_mode_selected():
            preview = self._auto_fan_percent()
            if preview is not None:
                return max(0, int(preview))
        if hasattr(self, 'fan_slider'):
            return max(0, int(self.fan_slider.value()))
        if self.settings.fan_is_off:
            return 0
        return self._duty_to_fan_slider(self.settings.current_fan_speed)

    def _current_export_pump_voltage(self):
        if self.pump_runtime_on and self.pump_runtime_voltage is not None:
            return pump_enum_to_display(self.pump_runtime_voltage)
        if self.auto_mode_active and self._auto_applied_pump_value is not None:
            return max(0, int(self._auto_applied_pump_value))
        # 自动模式已选中但尚未正式启用时，API 输出当前温度下的预估值，便于外部面板展示。
        if self._is_auto_mode_selected() and hasattr(self, 'pump_curve_widget'):
            temp = self._auto_control_temperature()
            if temp is not None:
                raw_fan = int(self.curve_widget.interpolate(temp)) if hasattr(self, 'curve_widget') else 0
                raw_pump = int(self.pump_curve_widget.interpolate(temp))
                _, pump_value = self._stabilize_auto_targets(temp, raw_fan, raw_pump)
                return max(0, int(pump_value))
        if hasattr(self, 'pump_slider'):
            voltage = self._pump_slider_to_voltage(self.pump_slider.value())
            return 0 if voltage is None else pump_enum_to_display(voltage)
        if self.settings.pump_is_off:
            return 0
        return pump_enum_to_display(self.settings.current_voltage)

    def _apply_export_api_settings(self, save=True):
        enabled = bool(getattr(self.settings, 'export_api_enabled', False))
        port = int(getattr(self.settings, 'export_api_port', DEFAULT_EXPORT_API_PORT))
        if self.export_api_server is not None:
            same_endpoint = self.export_api_server.host == DEFAULT_EXPORT_API_HOST and self.export_api_server.port == port
            if (not enabled) or (not same_endpoint):
                self.export_api_server.stop()
                self.export_api_server = None
        if enabled and self.export_api_server is None:
            self.export_api_server = WatercoolerApiServer(self.export_api_state, DEFAULT_EXPORT_API_HOST, port)
            self.export_api_server.start()
            if self.export_api_server._server is None:
                self.export_api_server = None
        if save:
            self.settings.save()
        self._update_export_api_controls()
        self._refresh_export_api_state()

    def _update_export_api_controls(self):
        if not hasattr(self, 'export_api_enable_checkbox'):
            return
        enabled = bool(getattr(self.settings, 'export_api_enabled', False))
        port = int(getattr(self.settings, 'export_api_port', DEFAULT_EXPORT_API_PORT))
        self.export_api_enable_checkbox.blockSignals(True)
        self.export_api_port_spin.blockSignals(True)
        try:
            if self.export_api_enable_checkbox.isChecked() != enabled:
                self.export_api_enable_checkbox.setChecked(enabled)
            self.export_api_port_spin.setEnabled(enabled)
            self.export_api_port_label.setEnabled(enabled)
            if self.export_api_port_spin.value() != port:
                self.export_api_port_spin.setValue(port)
        finally:
            self.export_api_enable_checkbox.blockSignals(False)
            self.export_api_port_spin.blockSignals(False)
        if hasattr(self, 'export_api_status_label'):
            if not enabled:
                self.export_api_status_label.setText('API 已关闭')
            elif self.export_api_server is None or self.export_api_server._server is None:
                self.export_api_status_label.setText(f'API 启动失败：127.0.0.1:{port}')
            else:
                self.export_api_status_label.setText(f'API：127.0.0.1:{port}')

    def _refresh_export_api_state(self):
        fan_percent = self._current_export_fan_percent()
        pump_voltage = self._current_export_pump_voltage()
        connected = bool(self.client and self.client.is_connected)
        device_name = None
        if hasattr(self, 'device_combo') and self.device_combo.count() > 0:
            device_name = self.device_combo.currentText()
        control_temp = self._current_control_temperature()
        api_enabled = bool(getattr(self.settings, 'export_api_enabled', False))
        api_port = int(getattr(self.settings, 'export_api_port', DEFAULT_EXPORT_API_PORT))
        api_running = bool(api_enabled and self.export_api_server is not None and self.export_api_server._server is not None)
        auto_selected = self._is_auto_mode_selected()
        self.export_api_state.update(
            connected=connected,
            mode='auto' if auto_selected else 'manual',
            auto={
                'selected': auto_selected,
                'active': bool(self.auto_mode_active),
            },
            device_name=device_name,
            fan={
                'percent': fan_percent,
                'text': f'{fan_percent}%',
                'is_off': fan_percent <= 0,
            },
            pump={
                'voltage': pump_voltage,
                'text': pump_curve_value_to_text(pump_voltage),
                'is_off': pump_voltage <= 0,
            },
            temperature={
                'cpu_c': None if self.last_cpu_temp is None else round(float(self.last_cpu_temp), 1),
                'gpu_c': None if self.last_gpu_temp is None else round(float(self.last_gpu_temp), 1),
                'control_c': None if control_temp is None else round(float(control_temp), 1),
            },
            api={
                'enabled': api_enabled,
                'running': api_running,
                'host': DEFAULT_EXPORT_API_HOST,
                'port': api_port,
                'status_url': f'http://{DEFAULT_EXPORT_API_HOST}:{api_port}/api/status' if api_enabled else None,
            },
        )

    def _fan_slider_to_duty(self, slider_value):
        slider_value = int(slider_value)
        if slider_value <= 0:
            return None
        return fan_percent_to_duty(slider_value)

    def _duty_to_fan_slider(self, duty):
        if duty is None:
            return 0
        try:
            duty = int(duty)
        except Exception:
            return 0
        if duty <= 0:
            return 0
        return clamp_curve_percent(round(duty / 255.0 * 100.0))

    def _pump_slider_to_voltage(self, slider_value):
        return {0: None, 1: PumpVoltage.V7, 2: PumpVoltage.V8, 3: PumpVoltage.V11}.get(slider_value)

    def _voltage_to_pump_slider(self, voltage):
        mapping = {1: PumpVoltage.V7, 2: PumpVoltage.V8, 3: PumpVoltage.V11}
        for idx, value in mapping.items():
            if value == voltage:
                return idx
        return 0

    async def _set_pump_runtime(self, should_run: bool, voltage: PumpVoltage | None = None):
        if not self.client or not self.client.is_connected:
            self.pump_runtime_on = False
            self.pump_runtime_voltage = None
            self._refresh_export_api_state()
            return
        if should_run and voltage is not None:
            await write_pump_mode(self.client, voltage)
            self.pump_runtime_on = True
            self.pump_runtime_voltage = voltage
        else:
            await write_pump_off(self.client)
            self.pump_runtime_on = False
            self.pump_runtime_voltage = None
        self._refresh_export_api_state()

    async def _apply_temperature_rgb_if_needed(self, temp=None, force=False):
        if not hasattr(self, 'rgb_temp_enabled_checkbox') or not self.rgb_temp_enabled_checkbox.isChecked():
            self._last_temp_rgb_bucket = None
            return
        if not self.client or not self.client.is_connected:
            return
        payload = self._temperature_rgb_payload(temp)
        if not force and payload['bucket'] == self._last_temp_rgb_bucket:
            return
        await write_rgb_mode(self.client, payload['mode'], payload['color'])
        self._last_temp_rgb_bucket = payload['bucket']

    async def _apply_auto_runtime(self, reason='temperature_tick'):
        if self.is_connecting or self.is_disconnecting or self._is_exiting:
            return
        if not self.client or not self.client.is_connected:
            self.auto_status_label.setText('自动模式待连接设备')
            return
        raw_temp = self._current_control_temperature()
        control_temp = self._auto_control_temperature()
        self._save_curve_settings_from_ui()
        self._save_rgb_temp_settings_from_ui()
        self._save_auto_debounce_settings_from_ui()
        if control_temp is None:
            self.auto_status_label.setText('自动模式已启用，但暂未读取到温度数据')
            return
        raw_fan_percent = int(self.curve_widget.interpolate(control_temp))
        raw_pump_value = int(self.pump_curve_widget.interpolate(control_temp))
        fan_percent, pump_value = self._stabilize_auto_targets(control_temp, raw_fan_percent, raw_pump_value)
        if fan_percent <= 0:
            await write_fan_off(self.client)
        else:
            await write_fan_mode(self.client, fan_percent_to_duty(fan_percent))
        if pump_value <= 0:
            await self._set_pump_runtime(False)
        else:
            await self._set_pump_runtime(True, pump_display_to_enum(pump_value))
        self._auto_applied_fan_percent = fan_percent
        self._auto_applied_pump_value = pump_value
        await self._apply_temperature_rgb_if_needed(raw_temp)
        temp_text = f"{int(control_temp)}°C"
        if raw_temp is not None and abs(float(raw_temp) - float(control_temp)) >= 0.5:
            temp_text = f"{int(control_temp)}°C（当前 {int(raw_temp)}°C）"
        self.auto_status_label.setText(
            f"自动模式运行中：{temp_text} → 风扇 {'关闭' if fan_percent <= 0 else f'{fan_percent}%'} / 水泵 {pump_curve_value_to_text(pump_value)}"
        )
        self.settings.auto_mode_enabled = True
        self.settings.selected_mode_index = 1
        self.settings.save()
        self._refresh_export_api_state()

    def sync_ui_from_settings(self):
        self._syncing_ui = True
        try:
            self.fan_slider.setValue(0 if self.settings.fan_is_off else self._duty_to_fan_slider(self.settings.current_fan_speed))
            self.pump_slider.setValue(0 if self.settings.pump_is_off else self._voltage_to_pump_slider(self.settings.current_voltage))

            self.fan_curve_points = normalize_fan_curve_points(self.settings.fan_curve_points)
            self.pump_curve_points = normalize_pump_curve_points(self.settings.pump_curve_points)
            self.curve_widget.points = [tuple(point) for point in self.fan_curve_points]
            self.pump_curve_widget.points = [tuple(point) for point in self.pump_curve_points]

            target_rgb_state = RGBMode.OFF if self.settings.rgb_is_off else self.settings.rgb_state
            mode_index = self.rgb_mode.findData(target_rgb_state)
            if mode_index >= 0:
                self.rgb_mode.setCurrentIndex(mode_index)
            self._set_combo_color_value(self.rgb_color, self.settings.rgb_color)

            self.rgb_temp_enabled_checkbox.setChecked(self.settings.rgb_temp_enabled)
            temp_mode_index = self.rgb_temp_mode_combo.findData(self.settings.rgb_temp_mode)
            if temp_mode_index >= 0:
                self.rgb_temp_mode_combo.setCurrentIndex(temp_mode_index)
            self.rgb_temp_low_spin.setValue(self.settings.rgb_temp_threshold_low)
            self.rgb_temp_high_spin.setValue(self.settings.rgb_temp_threshold_high)
            for combo, value in (
                (self.rgb_temp_low_color, self.settings.rgb_temp_color_low),
                (self.rgb_temp_mid_color, self.settings.rgb_temp_color_mid),
                (self.rgb_temp_high_color, self.settings.rgb_temp_color_high),
            ):
                self._set_combo_color_value(combo, value)

            self.auto_hysteresis_spin.setValue(int(self.settings.auto_hysteresis_c))
            samples_index = self.auto_samples_combo.findData(int(self.settings.auto_debounce_samples))
            if samples_index >= 0:
                self.auto_samples_combo.setCurrentIndex(samples_index)
            self.auto_fan_min_toggle_spin.setValue(float(self.settings.auto_fan_min_toggle_interval_sec))
            self.auto_pump_min_toggle_spin.setValue(float(self.settings.auto_pump_min_toggle_interval_sec))
            self._refresh_control_temp_history_maxlen()

            interval = normalize_update_interval(self.settings.update_interval_sec)
            preset_index = next((i for i, (sec, _) in enumerate(self.UPDATE_INTERVALS) if abs(sec - interval) < 0.001), -1)
            if preset_index >= 0:
                self.update_speed_combo.setCurrentIndex(preset_index)
                self.update_speed_custom.setEnabled(False)
                self.update_speed_custom.hide()
            else:
                self.update_speed_combo.setCurrentIndex(self.update_speed_combo.findData('custom'))
                self.update_speed_custom.setEnabled(True)
                self.update_speed_custom.show()
            self.update_speed_custom.setValue(interval)
            self.set_update_interval(interval, save=False)

            theme_index = self.theme_combo.findData(self.settings.theme_mode)
            if theme_index >= 0:
                self.theme_combo.setCurrentIndex(theme_index)

            self.export_api_enable_checkbox.setChecked(bool(self.settings.export_api_enabled))
            self.export_api_port_spin.setValue(int(self.settings.export_api_port))
            self._update_export_api_controls()

            self.mode_combo.setCurrentIndex(self.settings.selected_mode_index)
            self.auto_mode_active = bool(self.settings.auto_mode_enabled and self.settings.selected_mode_index == 1)
            self._reset_auto_control_state()
            self.auto_status_label.setText('自动模式已启用，等待连接或温度数据' if self.auto_mode_active else '自动模式未启用')
        finally:
            self._syncing_ui = False

        self._apply_theme(save=False)
        self._sync_rgb_input_states()
        self._update_mode_hint()
        self._update_control_summaries()
        self._update_manual_apply_hint()
        self._refresh_export_api_state()
        self._update_connection_controls()

    async def apply_saved_device_settings(self):
        if not self.client or not self.client.is_connected:
            return
        if self.settings.selected_mode_index == 1 and self.settings.auto_mode_enabled:
            self.auto_mode_active = True
            self._reset_auto_control_state()
            await self._apply_auto_runtime('connect')
        else:
            if self.settings.fan_is_off:
                await write_fan_off(self.client)
            else:
                duty = self.settings.current_fan_speed if self.settings.current_fan_speed is not None else 150
                await write_fan_mode(self.client, clamp_fan_duty(duty))
            if self.settings.pump_is_off:
                await self._set_pump_runtime(False)
            else:
                await self._set_pump_runtime(True, self.settings.current_voltage)
            self.auto_mode_active = False
            self.auto_status_label.setText('自动模式未启用')

        if self.settings.rgb_temp_enabled:
            await self._apply_temperature_rgb_if_needed(self._current_control_temperature(), force=True)
        elif self.settings.rgb_is_off or self.settings.rgb_state == RGBMode.OFF:
            await write_rgb_off(self.client)
        elif self.settings.rgb_state in [RGBMode.COLORFUL, RGBMode.BREATHE_COLOR]:
            await write_rgb_mode(self.client, self.settings.rgb_state, (0, 0, 0))
        else:
            await write_rgb_mode(self.client, self.settings.rgb_state, self.settings.rgb_color)
        self._update_control_summaries()

    def set_update_interval(self, sec, save=True):
        sec = normalize_update_interval(sec)
        if hasattr(self, 'temp_timer'):
            self.temp_timer.setInterval(int(sec * 1000))
        self.settings.update_interval_sec = sec
        if save:
            self.settings.save()

    def update_interval_changed(self, *args):
        if getattr(self, '_syncing_ui', False):
            return
        current = self.update_speed_combo.currentData()
        if current == 'custom':
            self.update_speed_custom.setEnabled(True)
            self.update_speed_custom.show()
            sec = self.update_speed_custom.value()
        else:
            self.update_speed_custom.setEnabled(False)
            self.update_speed_custom.hide()
            sec = current
        self.set_update_interval(sec)

    def _cleanup_before_exit(self):
        asyncio.ensure_future(self._shutdown_and_quit())

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.WindowStateChange and not self._is_exiting:
            if self.windowState() & QtCore.Qt.WindowMaximized:
                self._restore_maximized = True
            elif self.windowState() == QtCore.Qt.WindowNoState:
                self._restore_maximized = False
            if self.isMinimized():
                QtCore.QTimer.singleShot(0, self.hide)
        super().changeEvent(event)

    def closeEvent(self, event):
        if self._is_exiting:
            event.accept()
            return
        event.ignore()
        asyncio.ensure_future(self._shutdown_and_quit())

    def show_window(self):
        if self.settings.theme_mode == 'system':
            self._apply_theme(save=False)
        if self._restore_maximized:
            self.showMaximized()
        else:
            self.showNormal()
        self.raise_()
        self.activateWindow()

    def exit_app(self):
        asyncio.ensure_future(self._shutdown_and_quit())

    def on_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self.show_window()

    def on_mode_changed(self, index):
        self.pages.setCurrentIndex(index)
        if getattr(self, '_syncing_ui', False):
            self._update_mode_hint()
            self._update_control_summaries()
            return
        self.settings.selected_mode_index = index
        self._reset_auto_control_state()
        if index == 0:
            self.auto_mode_active = False
            self.settings.auto_mode_enabled = False
            self.auto_status_label.setText('已切回手动模式，自动托管已停止')
        else:
            self.auto_status_label.setText('自动模式已就绪，等待启用')
        self._update_mode_hint()
        self._update_control_summaries()
        self._update_manual_apply_hint()
        self.settings.save()

    def on_auto_connect_changed(self, state):
        self.settings.auto_connect = (state == QtCore.Qt.Checked)
        self.settings.save()

    def on_auto_start_changed(self, state):
        enabled = (state == QtCore.Qt.Checked)
        logging.info('Auto start checkbox changed: %s', enabled)
        self.settings.set_autostart(enabled)

    def on_export_api_settings_changed(self, *args):
        if getattr(self, '_syncing_ui', False):
            return
        self.settings.export_api_enabled = self.export_api_enable_checkbox.isChecked()
        self.settings.export_api_port = int(self.export_api_port_spin.value())
        self._apply_export_api_settings(save=True)

    def on_theme_changed(self, index):
        if getattr(self, '_syncing_ui', False):
            return
        self.settings.theme_mode = self.theme_combo.currentData()
        self._apply_theme(save=True)

    def on_dingtalk_enable_changed(self, *args):
        if getattr(self, '_syncing_ui', False):
            return
        enabled = bool(self.dingtalk_enable_checkbox.isChecked())
        if enabled:
            accepted = self.open_dingtalk_settings_dialog(triggered_by_enable=True)
            if accepted:
                self.settings.dingtalk_webhook_enabled = True
                self.settings.save()
            else:
                self.dingtalk_enable_checkbox.blockSignals(True)
                try:
                    self.dingtalk_enable_checkbox.setChecked(False)
                finally:
                    self.dingtalk_enable_checkbox.blockSignals(False)
                self.settings.dingtalk_webhook_enabled = False
                self.settings.save()
        else:
            self.settings.dingtalk_webhook_enabled = False
            self.settings.save()
        self._update_dingtalk_controls()

    def _update_dingtalk_controls(self):
        if not hasattr(self, 'dingtalk_enable_checkbox'):
            return
        enabled = bool(getattr(self.settings, 'dingtalk_webhook_enabled', False))
        self.dingtalk_enable_checkbox.blockSignals(True)
        try:
            self.dingtalk_enable_checkbox.setChecked(enabled)
        finally:
            self.dingtalk_enable_checkbox.blockSignals(False)
        if hasattr(self, 'dingtalk_settings_btn'):
            self.dingtalk_settings_btn.setText('设置')
            self.dingtalk_settings_btn.setToolTip('配置钉钉 Webhook、加签和测试推送')

    def _save_dingtalk_dialog_values(self):
        if not hasattr(self, '_dingtalk_dialog_webhook_edit'):
            return False
        webhook = self._dingtalk_dialog_webhook_edit.text().strip()
        secret = self._dingtalk_dialog_secret_edit.text().strip()
        if not webhook:
            QtWidgets.QMessageBox.warning(self, '钉钉设置', '请先填写 Webhook 地址。')
            return False
        self.settings.dingtalk_webhook_url = webhook
        self.settings.dingtalk_webhook_secret = secret
        self.settings.save()
        return True

    def _set_dingtalk_test_status(self, text, success=None):
        if not hasattr(self, '_dingtalk_dialog_test_status') or self._dingtalk_dialog_test_status is None:
            return
        self._dingtalk_dialog_test_status.setText(text)
        if success is True:
            self._dingtalk_dialog_test_status.setStyleSheet('color: #169c50;')
        elif success is False:
            self._dingtalk_dialog_test_status.setStyleSheet('color: #d43f3a;')
        else:
            self._dingtalk_dialog_test_status.setStyleSheet('')

    def _on_dingtalk_dialog_accept(self, dialog):
        if not self._save_dingtalk_dialog_values():
            return
        dialog.accept()

    def _build_dingtalk_settings_dialog(self):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('钉钉推送设置')
        dialog.setModal(True)
        dialog.resize(620, 240)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        tips_label = QtWidgets.QLabel('填写钉钉机器人 Webhook 与加签密钥。保存后，连接状态变化时会自动推送通知。')
        tips_label.setWordWrap(True)
        layout.addWidget(tips_label)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        form.setFormAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        webhook_edit = QtWidgets.QLineEdit()
        webhook_edit.setPlaceholderText('https://oapi.dingtalk.com/robot/send?access_token=...')
        webhook_edit.setText(str(getattr(self.settings, 'dingtalk_webhook_url', '') or ''))
        form.addRow('Webhook：', webhook_edit)

        secret_edit = QtWidgets.QLineEdit()
        secret_edit.setPlaceholderText('SEC...，未启用加签可留空')
        secret_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        secret_edit.setText(str(getattr(self.settings, 'dingtalk_webhook_secret', '') or ''))
        form.addRow('加签：', secret_edit)
        layout.addLayout(form)

        status_label = QtWidgets.QLabel('')
        status_label.setWordWrap(True)
        layout.addWidget(status_label)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.setSpacing(10)
        test_btn = QtWidgets.QPushButton('测试推送')
        save_btn = QtWidgets.QPushButton('保存')
        cancel_btn = QtWidgets.QPushButton('取消')
        btn_row.addWidget(test_btn)
        btn_row.addStretch()
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        self._dingtalk_settings_dialog = dialog
        self._dingtalk_dialog_webhook_edit = webhook_edit
        self._dingtalk_dialog_secret_edit = secret_edit
        self._dingtalk_dialog_test_btn = test_btn
        self._dingtalk_dialog_test_status = status_label

        test_btn.clicked.connect(self._test_dingtalk_push_from_dialog)
        save_btn.clicked.connect(lambda: self._on_dingtalk_dialog_accept(dialog))
        cancel_btn.clicked.connect(dialog.reject)
        webhook_edit.returnPressed.connect(lambda: self._on_dingtalk_dialog_accept(dialog))
        secret_edit.returnPressed.connect(lambda: self._on_dingtalk_dialog_accept(dialog))
        return dialog

    def open_dingtalk_settings_dialog(self, checked=False, triggered_by_enable=False):
        dialog = self._build_dingtalk_settings_dialog()
        self._set_dingtalk_test_status('', success=None)
        accepted = dialog.exec_() == QtWidgets.QDialog.Accepted
        self._dingtalk_settings_dialog = None
        self._dingtalk_dialog_webhook_edit = None
        self._dingtalk_dialog_secret_edit = None
        self._dingtalk_dialog_test_btn = None
        self._dingtalk_dialog_test_status = None
        if accepted:
            self._update_dingtalk_controls()
        return accepted

    def _test_dingtalk_push_from_dialog(self):
        if not hasattr(self, '_dingtalk_dialog_webhook_edit') or self._dingtalk_dialog_webhook_edit is None:
            return
        webhook = self._dingtalk_dialog_webhook_edit.text().strip()
        secret = self._dingtalk_dialog_secret_edit.text().strip()
        if not webhook:
            QtWidgets.QMessageBox.warning(self, '钉钉设置', '请先填写 Webhook 地址。')
            return
        webhook_url = self._build_dingtalk_webhook_url_from_values(webhook, secret)
        if not webhook_url:
            QtWidgets.QMessageBox.warning(self, '钉钉设置', 'Webhook 地址无效。')
            return
        if self._dingtalk_dialog_test_btn is not None:
            self._dingtalk_dialog_test_btn.setEnabled(False)
        self._set_dingtalk_test_status('正在发送测试推送...', success=None)
        message = self._format_dingtalk_message('水冷管理器通知', ['测试推送', f'当前模式：{self._current_mode_text()}'])

        def _worker():
            success, info = self._send_dingtalk_request(message, webhook_url)
            self.dingtalk_test_result.emit(success, info)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_dingtalk_test_result(self, success: bool, info: str):
        if getattr(self, '_dingtalk_dialog_test_btn', None) is not None:
            self._dingtalk_dialog_test_btn.setEnabled(True)
        if success:
            self._set_dingtalk_test_status('测试推送成功。', success=True)
        else:
            detail = info or '未知错误'
            self._set_dingtalk_test_status(f'测试推送失败：{detail}', success=False)

    def _current_mode_text(self):
        if hasattr(self, 'mode_combo') and self.mode_combo.currentIndex() == 1:
            return '自动模式'
        return '手动模式'

    def _build_dingtalk_webhook_url_from_values(self, webhook: str, secret: str = ''):
        webhook = str(webhook or '').strip()
        if not webhook:
            return None
        secret = str(secret or '').strip()
        if not secret:
            return webhook
        timestamp = str(int(time.time() * 1000))
        string_to_sign = f"{timestamp}\n{secret}"
        digest = hmac.new(secret.encode('utf-8'), string_to_sign.encode('utf-8'), digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(digest).decode('utf-8'))
        parsed = urllib.parse.urlsplit(webhook)
        query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
        query = [(k, v) for k, v in query if k not in ('timestamp', 'sign')]
        query.extend([('timestamp', timestamp), ('sign', sign)])
        return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), parsed.fragment))

    def _build_dingtalk_webhook_url(self):
        webhook = str(getattr(self.settings, 'dingtalk_webhook_url', '') or '').strip()
        secret = str(getattr(self.settings, 'dingtalk_webhook_secret', '') or '').strip()
        return self._build_dingtalk_webhook_url_from_values(webhook, secret)

    def _send_dingtalk_request(self, content: str, webhook_url: str):
        try:
            data = json.dumps({'msgtype': 'text', 'text': {'content': content}}, ensure_ascii=False).encode('utf-8')
            req = urllib.request.Request(webhook_url, data=data, headers={'Content-Type': 'application/json; charset=utf-8'}, method='POST')
            with urllib.request.urlopen(req, timeout=5) as resp:
                body = resp.read().decode('utf-8', errors='ignore')
                status_code = getattr(resp, 'status', 200)
            logging.info('DingTalk push response: %s', body[:400])
            if status_code >= 400:
                return False, f'HTTP {status_code}'
            try:
                payload = json.loads(body)
                if payload.get('errcode') not in (0, '0', None):
                    return False, payload.get('errmsg') or body[:200]
            except Exception:
                pass
            return True, body[:200]
        except Exception as exc:
            logging.exception('DingTalk push failed')
            return False, str(exc)

    def _send_dingtalk_text(self, content: str):
        if not bool(getattr(self.settings, 'dingtalk_webhook_enabled', False)):
            return
        webhook_url = self._build_dingtalk_webhook_url()
        if not webhook_url:
            logging.info('Skip DingTalk push: webhook not configured')
            return

        def _worker():
            self._send_dingtalk_request(content, webhook_url)

        threading.Thread(target=_worker, daemon=True).start()

    def _format_dingtalk_message(self, title: str, lines):
        now_text = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        content_lines = [str(title).strip(), f'时间：{now_text}']
        content_lines.extend([str(line).strip() for line in lines if str(line).strip()])
        return '\n'.join(content_lines)

    def _notify_connection_event(self, event: str, device_name: str = None):
        try:
            device_text = (device_name or self.selected_device_label.text().replace('设备：', '').strip() or self._last_known_device_name())
            mode_text = self._current_mode_text()
            if event == 'connected':
                message = self._format_dingtalk_message('水冷管理器通知', [f'已连接至 {device_text}', f'当前模式：{mode_text}'])
            elif event == 'auto_reconnected':
                message = self._format_dingtalk_message('水冷管理器通知', [f'已自动重连至 {device_text}', f'当前模式：{mode_text}'])
            elif event == 'disconnected':
                message = self._format_dingtalk_message('水冷管理器通知', ['已断开连接'])
            elif event == 'unexpected_disconnect':
                message = self._format_dingtalk_message('水冷管理器通知', [f'意外断开：{device_text}'])
            else:
                return
            self._send_dingtalk_text(message)
        except Exception:
            logging.exception('Prepare DingTalk notification failed: event=%s', event)

    def trigger_rescan(self):
        if self.is_scanning or self.is_connecting or self.is_disconnecting or self._is_exiting:
            return
        self.device_tip_label.setText("正在重新扫描设备，请稍候…")
        self._set_status_text("正在重新扫描设备...", connected=False)
        self._update_connection_controls()
        asyncio.ensure_future(self.scan_and_populate())

    @asyncSlot()
    async def scan_and_populate(self):
        if self.is_scanning:
            logging.info('Skip BLE scan: already scanning')
            return
        self.is_scanning = True
        logging.info('scan_and_populate started')
        self._update_connection_controls()
        try:
            devices = await scan_devices()
            self.device_combo.clear()
            for n, a in devices:
                self.device_combo.addItem(f"{n} [{a}]", a)
            if devices:
                logging.info('scan_and_populate found devices=%s', devices)
                self.device_combo.setEnabled(True)
                self.selected_device_label.setText(f"设备：{self.device_combo.currentText()}")
                self.device_tip_label.setText("已发现设备，选择后点击连接。")
                self._set_status_text(UI['select_prompt'], connected=False)
                self._refresh_export_api_state()
                self._update_connection_controls()
                if self.settings.auto_connect and not self.client:
                    await self.connect_device(auto_selected=True)
            else:
                logging.info('scan_and_populate found no live BLE device; fallback to last-known if available')
                last_addr = self._last_known_device_address()
                last_name = self._last_known_device_name()
                if last_addr:
                    self.device_combo.addItem(f"{last_name} [{last_addr}]（上次连接）", last_addr)
                    self.device_combo.setEnabled(True)
                    self.selected_device_label.setText(f"设备：{last_name} [{last_addr}]（上次连接）")
                    self.device_tip_label.setText("本次扫描未发现设备，但可尝试直连上次连接的设备。程序会继续自动重试扫描。")
                    self._set_status_text("扫描未发现设备，可尝试直连上次设备", connected=False)
                else:
                    self.device_combo.setEnabled(False)
                    self.selected_device_label.setText("设备：未发现")
                    self.device_tip_label.setText("暂未扫描到受支持的水冷设备，程序会自动重试。")
                    self._set_status_text(UI['no_device'], connected=False)
                self._refresh_export_api_state()
                self._update_connection_controls()
                QtCore.QTimer.singleShot(5000, lambda: asyncio.ensure_future(self.scan_and_populate()))
        finally:
            self.is_scanning = False
            logging.info('scan_and_populate finished')
            self._update_connection_controls()

    @asyncSlot()
    async def connect_device(self, auto_selected=False):
        if self.is_connecting or self.is_disconnecting or self._is_exiting:
            return
        if self.client and self.client.is_connected:
            self._update_connection_controls()
            return

        addr = self.device_combo.currentData()
        if not addr:
            self._set_status_text(UI['select_prompt'], connected=False)
            self._update_connection_controls()
            return

        self.is_connecting = True
        self._update_connection_controls()
        self.selected_device_label.setText(f"设备：{self.device_combo.currentText()}")
        self.device_tip_label.setText("正在建立蓝牙连接，请稍候…")
        self._set_status_text(UI['connecting'].format(addr), connected=False)
        def _on_client_disconnected(_client):
            logging.info('Bleak disconnected callback fired for %s', addr)
            self._queue_disconnect_handler(reason='bleak_callback')

        client = BleakClient(addr, disconnected_callback=_on_client_disconnected)
        try:
            try:
                if hasattr(client, 'set_disconnected_callback'):
                    client.set_disconnected_callback(_on_client_disconnected)
            except Exception:
                logging.exception('Failed to register BLE disconnected callback')
            await client.connect(timeout=8.0)
            if client.is_connected:
                self.client = client
                self.settings.last_device_address = addr
                self.settings.last_device_name = self.device_combo.currentText()
                self.settings.save()
                logging.info('BLE connected: %s', self.device_combo.currentText())
                self.device_tip_label.setText("设备已连接，可以直接应用设置。")
                self._set_status_text(UI['connected'].format(self.device_combo.currentText()), connected=True)
                self._set_manual_action_buttons_enabled(True)
                await self.apply_saved_device_settings()
                if auto_selected and self._auto_reconnect_pending:
                    self._notify_connection_event('auto_reconnected', self.device_combo.currentText())
                    self._auto_reconnect_pending = False
                else:
                    self._notify_connection_event('connected', self.device_combo.currentText())
                    self._auto_reconnect_pending = False
                self._refresh_export_api_state()
        except Exception as exc:
            logging.exception('BLE connect failed: %s', addr)
            self.client = None
            exc_name = exc.__class__.__name__
            if exc_name == 'BleakDeviceNotFoundError' or 'not found' in str(exc).lower():
                self.device_tip_label.setText("直连失败：当前未发现该设备广播，程序将自动重新扫描。")
                self._set_status_text("设备当前未发现，正在重新扫描", connected=False)
                if not self.is_scanning:
                    QtCore.QTimer.singleShot(800, lambda: asyncio.ensure_future(self.scan_and_populate()))
            else:
                self.device_tip_label.setText("连接失败，请重试；若扫描不到，也可继续尝试连接上次设备。")
                self._set_status_text("连接失败", connected=False)
        finally:
            self.is_connecting = False
            self._update_connection_controls()
            self._refresh_export_api_state()

    @asyncSlot()
    async def disconnect_device(self):
        if self.is_connecting or self.is_disconnecting or self._is_exiting:
            return
        if not self.client or not self.client.is_connected:
            self._update_connection_controls()
            return

        self.is_disconnecting = True
        self._update_connection_controls()
        self.device_tip_label.setText("正在断开设备…")
        self._set_status_text("正在断开设备...", connected=True)
        try:
            await self._disconnect_client(send_reset=True, update_status=True)
            self._auto_reconnect_pending = False
            self._notify_connection_event('disconnected')
        finally:
            self.is_disconnecting = False
            self._update_connection_controls()
            self._refresh_export_api_state()

    @asyncSlot()
    async def update_temperatures(self):
        if self._temperature_update_in_progress:
            return
        self._temperature_update_in_progress = True
        try:
            loop = asyncio.get_running_loop()
            try:
                cpu, gpu = await loop.run_in_executor(None, get_temperatures)
            except Exception:
                logging.exception('Temperature read failed')
                cpu = gpu = None
            if cpu is not None:
                self.last_cpu_temp = cpu
            if gpu is not None:
                self.last_gpu_temp = gpu
            cpu_txt = f"{int(self.last_cpu_temp)} °C" if self.last_cpu_temp is not None else "-- °C"
            gpu_txt = f"{int(self.last_gpu_temp)} °C" if self.last_gpu_temp is not None else "-- °C"
            self.cpu_value_label.setText(cpu_txt)
            self.gpu_value_label.setText(gpu_txt)
            self._update_control_temperature_history(self._current_control_temperature())
            self._update_control_summaries()

            if self.is_connecting or self.is_disconnecting or self._is_exiting:
                return

            try:
                if self.auto_mode_active and self.client and self.client.is_connected:
                    await self._apply_auto_runtime('temperature_tick')
                elif self.client and self.client.is_connected and self.rgb_temp_enabled_checkbox.isChecked():
                    await self._apply_temperature_rgb_if_needed(self._current_control_temperature())
            except Exception as exc:
                logging.exception('Temperature tick runtime apply failed')
                await self._handle_unexpected_disconnect(reason=exc)
        finally:
            self._temperature_update_in_progress = False
            self._refresh_export_api_state()

    @asyncSlot()
    async def apply_fan_and_pump(self):
        if not self.client or not self.client.is_connected:
            return
        try:
            self.auto_mode_active = False
            self._reset_auto_control_state()
            self.settings.auto_mode_enabled = False

            fan_index = self.fan_slider.value()
            duty = self._fan_slider_to_duty(fan_index)
            if duty is None:
                await write_fan_off(self.client)
                self.settings.fan_is_off = True
            else:
                safe_duty = clamp_fan_duty(duty)
                await write_fan_mode(self.client, safe_duty)
                self.settings.fan_is_off = False
                self.settings.current_fan_speed = safe_duty

            target_voltage = self._pump_slider_to_voltage(self.pump_slider.value())
            if target_voltage is None:
                await self._set_pump_runtime(False)
                self.settings.pump_is_off = True
            else:
                await self._set_pump_runtime(True, target_voltage)
                self.settings.pump_is_off = False
                self.settings.current_voltage = target_voltage

            self.settings.selected_mode_index = self.mode_combo.currentIndex()
            self._save_auto_debounce_settings_from_ui()
            self.settings.save()
            self.auto_status_label.setText('自动模式未启用')
            self._update_control_summaries()
            self._update_manual_apply_hint()
            self._refresh_export_api_state()
        except Exception as exc:
            logging.exception('Apply fan/pump failed')
            await self._handle_unexpected_disconnect(reason=exc)

    @asyncSlot()
    async def apply_rgb(self):
        if not self.client or not self.client.is_connected:
            return
        try:
            self._save_rgb_temp_settings_from_ui()
            if self.rgb_temp_enabled_checkbox.isChecked():
                self.settings.rgb_is_off = False
                self._last_temp_rgb_bucket = None
                await self._apply_temperature_rgb_if_needed(self._current_control_temperature(), force=True)
            else:
                mode = self.rgb_mode.currentData()
                color = self._combo_color_value(self.rgb_color, self.settings.rgb_color)
                if mode == RGBMode.OFF:
                    await write_rgb_off(self.client)
                    self.settings.rgb_is_off = True
                elif mode in [RGBMode.COLORFUL, RGBMode.BREATHE_COLOR]:
                    await write_rgb_mode(self.client, mode, (0, 0, 0))
                    self.settings.rgb_is_off = False
                else:
                    await write_rgb_mode(self.client, mode, color)
                    self.settings.rgb_is_off = False
                self.settings.rgb_state = mode
                self.settings.rgb_color = color
                self._last_temp_rgb_bucket = None
            self.settings.save()
            self._update_control_summaries()
            self._refresh_export_api_state()
        except Exception as exc:
            logging.exception('Apply RGB failed')
            await self._handle_unexpected_disconnect(reason=exc)

    @asyncSlot()
    async def apply_all(self):
        await self.apply_fan_and_pump()
        await self.apply_rgb()

    @asyncSlot()
    async def apply_curve(self):
        if not self.client or not self.client.is_connected:
            return
        try:
            self.auto_mode_active = True
            self._save_curve_settings_from_ui()
            self._save_rgb_temp_settings_from_ui()
            self._save_auto_debounce_settings_from_ui()
            self.settings.auto_mode_enabled = True
            self.settings.selected_mode_index = 1
            self._reset_auto_control_state()
            await self._apply_auto_runtime('manual_enable')
            self._update_control_summaries()
            self._refresh_export_api_state()
        except Exception as exc:
            logging.exception('Enable auto curve failed')
            await self._handle_unexpected_disconnect(reason=exc)


def main():
    setup_logging()
    if not ensure_admin_rights():
        return

    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    install_global_exception_hooks(loop)
    window = MainWindow()
    app.aboutToQuit.connect(lambda: logging.info('Qt aboutToQuit fired'))
    atexit.register(lambda: logging.info('Process exiting'))
    window.showMaximized()
    asyncio.ensure_future(window.scan_and_populate())
    with loop:
        loop.run_forever()

if __name__ == '__main__':
    main()

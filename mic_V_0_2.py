import json
import os
import threading
import time
import keyboard
import tkinter as tk
from tkinter import simpledialog, messagebox
import sys

import comtypes
from comtypes import CLSCTX_ALL, CoInitialize
from ctypes import cast, POINTER
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# 작업 트레이 아이콘 관련 라이브러리 (pystray, PIL)
import pystray
from PIL import Image, ImageDraw

# 언어별 메뉴 문자열 사전
LANG = {
    "Korean": {
        "mute": "눌렀을때 음소거",
        "unmute": "눌렀을때 음소거 해제",
        "volume_settings": "볼륨 설정 변경",
        "key_settings": "키 설정 변경",
        "language_settings": "언어설정",
        "korean": "한국어",
        "english": "English",
        "info": "정보",
        "exit": "종료"
    },
    "English": {
        "mute": "Mute when pressed",
        "unmute": "Unmute when pressed",
        "volume_settings": "Change volume settings",
        "key_settings": "Change key settings",
        "language_settings": "Language Settings",
        "korean": "Korean",
        "english": "English",
        "info": "Info",
        "exit": "Exit"
    }
}

# 기본 설정값
settings = {
    "volume": 50,                # 복원할 기본 마이크 볼륨 (0~100)
    "trigger_key": "ctrl+shift+m"  # 음소거/복원 단축키
}

# 전역 변수: 마이크 컨트롤러와 핫키 핸들
mic_controller = None
hotkey_press_hook = None
hotkey_release_hook = None

# 전역 변수: 토글 모드
# True: "눌렀을때 음소거" 상태, False: "눌렀을때 음소거 해제" 상태
always_on = True

# 전역 변수: 언어 (기본은 English)
language = "English"

# 전역 변수: Tk 다이얼로그 루트 (한번만 생성)
dialog_root = None

def init_dialog_root():
    global dialog_root
    if dialog_root is None:
        dialog_root = tk.Tk()
        dialog_root.withdraw()

# AppData에 설정 파일을 저장하는 함수
def get_config_path():
    appdata = os.getenv("APPDATA")
    config_dir = os.path.join(appdata, "MyMicController")
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    return os.path.join(config_dir, "config.json")

def load_settings():
    global settings, language
    config_path = get_config_path()
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            settings = data.get("settings", settings)
            language = data.get("language", language)
    else:
        create_settings_dialog()

def save_settings():
    config_path = get_config_path()
    data = {
        "settings": settings,
        "language": language
    }
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def create_settings_dialog():
    global settings, language
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("설정", "설정 파일이 존재하지 않습니다. 설정을 진행합니다.")
    vol = simpledialog.askinteger("마이크 볼륨 설정", "기본 마이크 볼륨을 입력하세요 (0~100):", minvalue=0, maxvalue=100)
    key = simpledialog.askstring("키 설정", "마이크 음소거/복원 단축키를 입력하세요 (예: ctrl+shift+m):")
    settings["volume"] = vol if vol is not None else 50
    settings["trigger_key"] = key if key else "ctrl+shift+m"
    language = "English"  # 기본 언어는 English
    save_settings()
    root.destroy()

class MicController:
    def __init__(self):
        self.device = self.get_default_microphone()
        self.interface = None
        self.get_interface()
    
    def get_default_microphone(self):
        CoInitialize()
        device = AudioUtilities.GetMicrophone()
        return device

    def get_interface(self):
        if self.device is not None:
            try:
                interface = self.device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                self.interface = cast(interface, POINTER(IAudioEndpointVolume))
            except Exception as e:
                print("마이크 인터페이스 활성화 실패:", e)
    
    def set_volume(self, volume_percent):
        if self.interface:
            vol_level = volume_percent / 100.0
            try:
                self.interface.SetMasterVolumeLevelScalar(vol_level, None)
            except Exception as e:
                print("볼륨 설정 실패:", e)
    
    def refresh_device(self):
        new_device = self.get_default_microphone()
        if new_device is not None and new_device != self.device:
            print("기본 마이크 변경 감지, 새로운 마이크 적용")
            self.device = new_device
            self.get_interface()

def on_trigger_press(event):
    global always_on
    if always_on:
        print("단축키 눌림: (눌렀을때 음소거 상태) 볼륨 0")
        mic_controller.set_volume(0)
    else:
        print("단축키 눌림: (눌렀을때 음소거 해제 상태) 최대 볼륨 100")
        mic_controller.set_volume(100)

def on_trigger_release(event):
    global always_on
    if always_on:
        print("단축키 해제: (눌렀을때 음소거 상태) 볼륨 복원")
        mic_controller.set_volume(settings["volume"])
    else:
        print("단축키 해제: (눌렀을때 음소거 해제 상태) 볼륨 0")
        mic_controller.set_volume(0)

def update_hotkeys():
    global hotkey_press_hook, hotkey_release_hook
    if hotkey_press_hook is not None:
        try:
            keyboard.unhook(hotkey_press_hook)
        except Exception as e:
            print("hotkey_press unhook 에러:", e)
    if hotkey_release_hook is not None:
        try:
            keyboard.unhook(hotkey_release_hook)
        except Exception as e:
            print("hotkey_release unhook 에러:", e)
    trigger_key = settings["trigger_key"]
    hotkey_press_hook = keyboard.on_press_key(trigger_key, on_trigger_press, suppress=False)
    hotkey_release_hook = keyboard.on_release_key(trigger_key, on_trigger_release, suppress=False)
    print(f"단축키 '{trigger_key}'로 등록되었습니다.")

def keyboard_thread():
    update_hotkeys()
    while True:
        time.sleep(1)

def monitor_default_mic():
    while True:
        mic_controller.refresh_device()
        time.sleep(5)

def create_image():
    image = Image.new('RGB', (64, 64), color='white')
    dc = ImageDraw.Draw(image)
    dc.ellipse((8, 8, 56, 56), fill='black')
    return image

# --- ask_for_key() 함수: 전역 dialog_root를 부모로 사용 ---
def ask_for_key():
    result = []

    def on_key(event):
        key = event.keysym
        keycode = event.keycode  # 키 코드도 저장
        print(f"입력된 키: {key}, Keycode: {keycode}")

        # Ctrl, Alt, Shift 등도 감지할 수 있도록 보정
        if key in ["Control_L", "Control_R", "Alt_L", "Alt_R", "Shift_L", "Shift_R", "Caps_Lock", "Tab"]:
            result.append(key)
        else:
            result.append(key)

        top.destroy()

    root = tk.Tk()
    root.withdraw()
    top = tk.Toplevel(root)
    top.title("키 설정 변경")
    label = tk.Label(top, text="원하는 키를 누르세요 (예: CapsLock, Tab, a, 등)")
    label.pack(padx=20, pady=20)
    
    top.bind("<KeyPress>", on_key)  # ✅ KeyPress 이벤트에서 감지
    top.focus_force()
    top.grab_set()
    top.wait_window()
    root.destroy()

    return result[0] if result else None


def on_info(icon, item):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Info", "producer : https://github.com/jaeminsd")
    root.destroy()

def set_always_on(icon, item):
    global always_on
    always_on = True
    print("항상 모드: 눌렀을때 음소거")
    icon.menu = setup_tray_menu()
    icon.update_menu()

def set_always_off(icon, item):
    global always_on
    always_on = False
    print("항상 모드: 눌렀을때 음소거 해제")
    icon.menu = setup_tray_menu()
    icon.update_menu()

def on_volume_settings(icon, item):
    def dialog():
        root = tk.Tk()
        root.withdraw()
        vol = simpledialog.askinteger(LANG[language]["volume_settings"], LANG[language]["volume_settings"], minvalue=0, maxvalue=100)
        if vol is not None:
            settings["volume"] = vol
            save_settings()
            print(f"기본 볼륨을 {vol}%로 변경하였습니다.")
        root.destroy()
    threading.Thread(target=dialog).start()

def on_key_settings(icon, item):
    key = ask_for_key()  # ✅ 메인 스레드에서 실행
    if key:
        settings["trigger_key"] = key
        save_settings()
        update_hotkeys()
        print(f"단축키를 '{key}'로 변경하였습니다.")


def set_language_korean(icon, item):
    global language
    language = "Korean"
    print("언어 설정: 한국어")
    icon.menu = setup_tray_menu()
    icon.update_menu()
    save_settings()

def set_language_english(icon, item):
    global language
    language = "English"
    print("Language set: English")
    icon.menu = setup_tray_menu()
    icon.update_menu()
    save_settings()

def on_exit(icon, item):
    icon.stop()
    os._exit(0)

def setup_tray_menu():
    lang = LANG[language]
    return pystray.Menu(
        pystray.MenuItem(lang["mute"], set_always_on, checked=lambda item: always_on),
        pystray.MenuItem(lang["unmute"], set_always_off, checked=lambda item: not always_on),
        pystray.MenuItem(lang["volume_settings"], on_volume_settings),
        pystray.MenuItem(lang["key_settings"], on_key_settings),
        pystray.MenuItem(lang["language_settings"], pystray.Menu(
            pystray.MenuItem(lang["korean"], set_language_korean, checked=lambda item: language=="Korean"),
            pystray.MenuItem(lang["english"], set_language_english, checked=lambda item: language=="English")
        )),
        pystray.MenuItem(lang["info"], on_info),
        pystray.MenuItem(lang["exit"], on_exit)
    )

def setup_tray_icon():
    menu = setup_tray_menu()
    icon = pystray.Icon("mic_controller", create_image(), "Mic Controller", menu)
    return icon

def main():
    global mic_controller, dialog_root
    load_settings()
    init_dialog_root()
    mic_controller = MicController()
    
    kb_thread = threading.Thread(target=keyboard_thread, daemon=True)
    kb_thread.start()
    
    monitor_thread = threading.Thread(target=monitor_default_mic, daemon=True)
    monitor_thread.start()
    
    tray_icon = setup_tray_icon()
    tray_icon.run()

if __name__ == "__main__":
    from PIL import Image, ImageDraw
    main()

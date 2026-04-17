import cv2
import tkinter as tk
from tkinter import messagebox, ttk
from pynput import mouse, keyboard
import threading
import time
import os
import pickle
from PIL import Image, ImageTk
import numpy as np
import winreg
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage
import json
import datetime


class PrivacyShield:
    def __init__(self):
        # 初始化日志系统
        import logging
        logging.basicConfig(
            filename='privacy_shield.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger('PrivacyShield')

        self.screenshot_dir = r'C:\Users\mahaoxing\Pictures\pricay'
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)
            self.logger.info(f"创建目录: {self.screenshot_dir}")

        self.config_file = 'privacy_shield_config.json'
        self.config = self.load_config()

        self.owner_file = 'owner_info.pkl'

        self.is_running = False
        self.is_black_screen = False
        self.stranger_start_time = None
        self.click_count = 0
        self.last_click_time = 0
        self.black_screen_window = None
        self.owner_detected = False
        self.owner_encodings = []
        self.tray_icon = None
        self.hotkey_listener = None
        self.show_camera = self.config.get('show_camera', False)
        self.camera_index = self.config.get('camera_index', 0)
        self.root = None

        self.video_capture = None
        self.listener = None

        # 延迟加载人脸检测器，加快启动速度
        self.face_cascade = None

        self.owner_registered = False
        self.need_config_wizard = False

        # 窗口状态标志
        self.config_wizard_window = None
        self.settings_window = None

        # 摄像头检测缓存
        self.camera_detection_cache = None
        self.camera_detection_time = 0
        if os.path.exists(self.owner_file):
            try:
                self.load_owner_info()
                self.owner_registered = True
                self.logger.info(f"已加载主人信息，包含 {len(self.owner_encodings)} 个角度的特征")
            except Exception as e:
                self.logger.error(f"加载主人信息失败: {e}")
                os.remove(self.owner_file)
                self.need_config_wizard = True
        else:
            self.need_config_wizard = True

        # 延迟初始化Tk根窗口，加快启动速度
        self.root = None

    def load_config(self):
        default_config = {
            'sensitivity': 0.45,
            'response_time': 0.5,
            'screenshot_path': r'C:\Users\mahaoxing\Pictures\pricay',
            'startup': False,
            'hotkey': 'ctrl+alt+p',
            'email_notification': False,
            'email': '',
            'smtp_server': '',
            'smtp_port': 587,
            'smtp_username': '',
            'smtp_password': '',
            'detection_area': [0, 0, 100, 100],
            'schedule': [],
            'show_camera': True,
            'camera_index': 0
        }

        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
            except:
                return default_config
        return default_config

    def save_config(self):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)

    def load_owner_info(self):
        with open(self.owner_file, 'rb') as f:
            self.owner_encodings = pickle.load(f)
        print(f"已加载主人信息，包含 {len(self.owner_encodings)} 个角度的特征")

    def save_owner_info(self):
        with open(self.owner_file, 'wb') as f:
            pickle.dump(self.owner_encodings, f)

    def show_config_wizard(self):
        # 确保主窗口存在
        self.create_root()
        if not self.root or not self.root.winfo_exists():
            self.logger.error("主窗口不存在，无法创建配置向导窗口")
            return

        # 检查配置向导窗口是否已经存在
        if self.config_wizard_window and self.config_wizard_window.winfo_exists():
            # 窗口已存在，将其置于前台
            self.config_wizard_window.deiconify()
            self.config_wizard_window.lift()
            self.config_wizard_window.focus_force()
            return

        # 临时保存并禁用 show_camera，避免自动弹出摄像头预览
        original_show_camera = self.show_camera
        self.show_camera = False

        # 直接创建窗口，不使用 after，确保在当前线程中执行
        try:
            self.logger.info("正在创建配置向导窗口")
            wizard = tk.Toplevel(self.root)
            wizard.title("防窥屏程序 - 配置向导")
            wizard.geometry("500x450")
            wizard.resizable(False, False)
            wizard.transient(self.root)
            # 不使用grab_set()，允许用户与多个窗口交互
            # 确保窗口可见
            wizard.deiconify()
            wizard.lift()
            wizard.focus_force()

            # 保存窗口引用
            self.config_wizard_window = wizard

            # 设置窗口关闭回调
            def on_wizard_close():
                self.config_wizard_window = None
                # 恢复原来的 show_camera 状态
                self.show_camera = original_show_camera
                wizard.destroy()

            wizard.protocol("WM_DELETE_WINDOW", on_wizard_close)
            self.logger.info("配置向导窗口已创建")

            steps = ["欢迎", "注册主人", "摄像头设置", "配置设置", "完成"]
            # 如果主人已注册，跳过注册步骤
            if self.owner_registered:
                current_step = 2  # 直接跳到摄像头设置
            else:
                current_step = 0

            def next_step():
                nonlocal current_step
                if current_step == 0:
                    current_step = 1
                    update_wizard()
                elif current_step == 1:
                    if not self.owner_registered:
                        if self.register_owner_wizard(wizard):
                            current_step = 2
                            update_wizard()
                    else:
                        current_step = 2
                        update_wizard()
                elif current_step == 2:
                    current_step = 3
                    update_wizard()
                elif current_step == 3:
                    save_settings()
                    current_step = 4
                    update_wizard()
                elif current_step == 4:
                    self.config_wizard_window = None
                    wizard.destroy()

            def prev_step():
                nonlocal current_step
                if current_step > 0:
                    current_step -= 1
                    update_wizard()

            def save_settings():
                try:
                    self.config['sensitivity'] = sensitivity_var.get()
                    self.config['response_time'] = response_var.get()
                    self.config['startup'] = startup_var.get()
                    # 保持 show_camera 为 True
                    self.config['show_camera'] = True
                    self.show_camera = True
                    self.save_config()
                    if startup_var.get():
                        self.set_startup(True)
                    else:
                        self.set_startup(False)
                except Exception as e:
                    self.logger.error(f"保存配置向导设置时出错: {e}")

            def update_wizard():
                for widget in wizard.winfo_children():
                    widget.destroy()

                ttk.Label(wizard, text=f"步骤 {current_step + 1}/{len(steps)}: {steps[current_step]}",
                          font=("Arial", 16, "bold")).pack(pady=20)

                if current_step == 0:
                    ttk.Label(wizard, text="欢迎使用防窥屏程序！", font=("Arial", 14)).pack(pady=10)
                    ttk.Label(wizard, text="此向导将帮助你配置防窥屏程序，保护你的隐私安全。",
                              wraplength=400).pack(pady=10)
                    ttk.Label(wizard, text="点击下一步开始配置。", wraplength=400).pack(pady=10)

                elif current_step == 1:
                    if self.owner_registered:
                        ttk.Label(wizard, text="主人信息已注册，点击下一步继续。",
                                  wraplength=400).pack(pady=10)
                    else:
                        ttk.Label(wizard, text="请面对摄像头，点击下方按钮开始注册你的面部信息。",
                                  wraplength=400).pack(pady=10)
                        ttk.Button(wizard, text="开始注册",
                                   command=lambda: self.register_owner_wizard(wizard)).pack(pady=10)

                elif current_step == 2:
                    ttk.Label(wizard, text="摄像头设置：", font=("Arial", 12, "bold")).pack(pady=10)

                    ttk.Label(wizard, text="程序将使用默认摄像头进行检测。", wraplength=400).pack(pady=10)
                    ttk.Label(wizard, text="默认开启实时摄像头检测画面。", wraplength=400).pack(pady=5)

                elif current_step == 3:
                    ttk.Label(wizard, text="请配置程序设置：", font=("Arial", 12, "bold")).pack(pady=10)

                    ttk.Label(wizard, text="识别灵敏度 (0.1-1.0)：").pack(pady=5)
                    sensitivity_var.set(self.config['sensitivity'])

                    sensitivity_frame = ttk.Frame(wizard)
                    sensitivity_frame.pack(pady=5)

                    sensitivity_scale = ttk.Scale(sensitivity_frame, from_=0.1, to=1.0, variable=sensitivity_var,
                                                  orient=tk.HORIZONTAL, length=300)
                    sensitivity_scale.pack(side=tk.LEFT, padx=10)

                    sensitivity_value = ttk.Label(sensitivity_frame, text="{:.2f}".format(sensitivity_var.get()),
                                                  width=10)
                    sensitivity_value.pack(side=tk.RIGHT, padx=10)

                    def update_sensitivity_value(*args):
                        sensitivity_value.config(text="{:.2f}".format(sensitivity_var.get()))

                    sensitivity_var.trace('w', update_sensitivity_value)

                    ttk.Label(wizard, text="灵敏度越高，识别越严格，可能会将主人误判为陌生人",
                              font=("Arial", 10), foreground="gray").pack(pady=2)

                    ttk.Label(wizard, text="响应时间 (秒)：").pack(pady=5)
                    response_var.set(self.config['response_time'])

                    response_frame = ttk.Frame(wizard)
                    response_frame.pack(pady=5)

                    response_scale = ttk.Scale(response_frame, from_=0.1, to=2.0, variable=response_var,
                                               orient=tk.HORIZONTAL, length=300)
                    response_scale.pack(side=tk.LEFT, padx=10)

                    response_value = ttk.Label(response_frame, text="{:.2f}".format(response_var.get()), width=10)
                    response_value.pack(side=tk.RIGHT, padx=10)

                    def update_response_value(*args):
                        response_value.config(text="{:.2f}".format(response_var.get()))

                    response_var.trace('w', update_response_value)

                    startup_var.set(self.config['startup'])
                    ttk.Checkbutton(wizard, text="开机自启动", variable=startup_var).pack(pady=10)

                elif current_step == 4:
                    ttk.Label(wizard, text="配置完成！", font=("Arial", 14)).pack(pady=10)
                    ttk.Label(wizard, text="防窥屏程序已成功配置，将开始保护你的隐私。",
                              wraplength=400).pack(pady=10)
                    ttk.Label(wizard, text="你可以通过系统托盘图标管理程序。",
                              wraplength=400).pack(pady=10)

                frame = ttk.Frame(wizard)
                frame.pack(pady=20)

                if current_step > 0:
                    ttk.Button(frame, text="上一步", command=prev_step).pack(side=tk.LEFT, padx=10)

                def on_complete():
                    self.config_wizard_window = None
                    self.show_camera = original_show_camera
                    wizard.destroy()

                if current_step < len(steps) - 1:
                    ttk.Button(frame, text="下一步", command=next_step).pack(side=tk.RIGHT, padx=10)
                else:
                    ttk.Button(frame, text="完成", command=on_complete).pack(side=tk.RIGHT, padx=10)

            sensitivity_var = tk.DoubleVar(value=self.config['sensitivity'])
            response_var = tk.DoubleVar(value=self.config['response_time'])
            startup_var = tk.BooleanVar(value=self.config['startup'])
            show_camera_var = tk.BooleanVar(value=self.config['show_camera'])

            update_wizard()
            wizard.wait_window()
        except Exception as e:
            self.logger.error(f"创建配置向导窗口时出错: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            try:
                messagebox.showerror("错误", f"创建配置向导窗口时出错: {e}\n请查看 privacy_shield.log 获取详细信息")
            except:
                pass

    def create_root(self):
        """创建主窗口"""
        if not self.root:
            self.root = tk.Tk()
            self.root.title("防窥屏程序")
            self.root.geometry("200x100")
            self.root.resizable(False, False)
            self.root.protocol("WM_DELETE_WINDOW", self.on_root_close)
            # 不隐藏主窗口，确保事件循环正常工作
            self.root.deiconify()
            # 最小化到任务栏，而不是完全隐藏
            self.root.iconify()
        # 确保主窗口是活跃的
        if self.root.winfo_exists():
            self.root.update_idletasks()

    def get_face_cascade(self):
        """获取人脸检测器，延迟加载"""
        if not self.face_cascade:
            try:
                self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
                self.logger.info("人脸检测器加载成功")
            except Exception as e:
                self.logger.error(f"加载人脸检测器失败: {e}")
                # 使用备用方法
                import sys
                cascade_path = os.path.join(os.path.dirname(sys.executable), 'Lib', 'site-packages', 'cv2', 'data',
                                            'haarcascade_frontalface_default.xml')
                if os.path.exists(cascade_path):
                    try:
                        self.face_cascade = cv2.CascadeClassifier(cascade_path)
                        self.logger.info(f"从备用路径加载人脸检测器成功: {cascade_path}")
                    except Exception as e2:
                        self.logger.error(f"备用路径加载失败: {e2}")
        return self.face_cascade

    def on_root_close(self):
        self.stop_protection()
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)

    def register_owner_wizard(self, parent=None):
        cap = None
        try:
            if parent:
                messagebox.showinfo("提示", "请面对摄像头，程序将从多个角度注册你的面部信息\n请按照提示做出不同的姿势",
                                    parent=parent)
            else:
                self.create_root()
                messagebox.showinfo("提示", "请面对摄像头，程序将从多个角度注册你的面部信息\n请按照提示做出不同的姿势")

            if self.video_capture:
                try:
                    self.video_capture.release()
                except Exception as e:
                    self.logger.error(f"释放摄像头资源时出错: {e}")
                self.video_capture = None

            # 尝试打开摄像头，使用多种后端
            cap = None
            backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF, 0]
            for backend in backends:
                try:
                    self.logger.info(f"尝试使用后端 {backend} 打开摄像头 {self.camera_index}")
                    cap = cv2.VideoCapture(self.camera_index, backend)
                    if cap and cap.isOpened():
                        self.logger.info(f"成功使用后端 {backend} 打开摄像头 {self.camera_index}")
                        break
                except Exception as e:
                    self.logger.warning(f"使用后端 {backend} 打开摄像头失败: {e}")
                finally:
                    if cap and not cap.isOpened():
                        try:
                            cap.release()
                        except:
                            pass
                        cap = None

            if not cap or not cap.isOpened():
                if parent:
                    messagebox.showerror("错误", "无法打开摄像头", parent=parent)
                else:
                    messagebox.showerror("错误", "无法打开摄像头")
                return False

            # 设置摄像头参数
            try:
                cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            except Exception as e:
                self.logger.warning(f"设置摄像头参数失败: {e}")

            # 验证是否能读取到帧
            ret, test_frame = cap.read()
            if not ret or test_frame is None or test_frame.size == 0:
                if parent:
                    messagebox.showerror("错误", "无法从摄像头获取画面", parent=parent)
                else:
                    messagebox.showerror("错误", "无法从摄像头获取画面")
                return False

            # 检查帧是否是全黑的
            try:
                gray = cv2.cvtColor(test_frame, cv2.COLOR_BGR2GRAY)
                mean_brightness = cv2.mean(gray)[0]
                if mean_brightness < 10:  # 全黑帧
                    if parent:
                        messagebox.showerror("错误", "摄像头画面黑屏，请检查摄像头连接", parent=parent)
                    else:
                        messagebox.showerror("错误", "摄像头画面黑屏，请检查摄像头连接")
                    return False
            except Exception as e:
                self.logger.warning(f"检查帧亮度失败: {e}")

            self.logger.info(f"成功打开摄像头 {self.camera_index}")

            # 获取人脸检测器
            face_cascade = self.get_face_cascade()
            if not face_cascade:
                if parent:
                    messagebox.showerror("错误", "无法加载人脸检测器", parent=parent)
                else:
                    messagebox.showerror("错误", "无法加载人脸检测器")
                return False

            owner_detected = 0
            self.owner_encodings = []

            poses = [
                "正面直视摄像头",
                "稍微向左转头",
                "稍微向右转头",
                "稍微抬头",
                "稍微低头",
                "微笑",
                "正常表情",
                "向左倾斜头部",
                "向右倾斜头部",
                "近距离看摄像头",
                "稍远距离看摄像头"
            ]

            for i, pose in enumerate(poses):
                for _ in range(20):
                    try:
                        ret, frame = cap.read()
                        if not ret:
                            time.sleep(0.01)
                            continue

                        if frame is None or frame.size == 0:
                            time.sleep(0.01)
                            continue

                        # 检查帧是否是全黑的
                        try:
                            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                            mean_brightness = cv2.mean(gray)[0]
                            if mean_brightness < 10:  # 全黑帧
                                time.sleep(0.01)
                                continue
                        except Exception as e:
                            self.logger.debug(f"检查帧亮度失败: {e}")

                        cv2.putText(frame, f"姿势 {i + 1}/{len(poses)}: {pose}", (10, 30),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

                        try:
                            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                            faces = face_cascade.detectMultiScale(
                                gray,
                                scaleFactor=1.1,
                                minNeighbors=5,
                                minSize=(100, 100)
                            )

                            if len(faces) == 1:
                                x, y, w, h = faces[0]
                                face_roi = gray[y:y + h, x:x + w]
                                face_roi = cv2.resize(face_roi, (100, 100))

                                hist = cv2.calcHist([face_roi], [0], None, [256], [0, 256])
                                hist = cv2.normalize(hist, hist).flatten()

                                self.owner_encodings.append(hist)
                                owner_detected += 1

                                if owner_detected % 5 == 0:
                                    cv2.putText(frame, f"已采集 {owner_detected} 个样本", (10, 70),
                                                cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
                        except Exception as e:
                            self.logger.debug(f"人脸检测错误: {e}")
                            continue

                        try:
                            cv2.imshow('注册主人面部', frame)
                            key = cv2.waitKey(100) & 0xFF
                            if key == ord('q'):
                                break
                        except Exception as e:
                            self.logger.warning(f"显示摄像头画面失败: {e}")
                            time.sleep(0.01)
                    except Exception as e:
                        self.logger.warning(f"采集样本时出错: {e}")
                        time.sleep(0.01)
                        continue

            if owner_detected >= 20:
                self.save_owner_info()
                self.owner_registered = True
                if parent:
                    messagebox.showinfo("成功", f"主人面部注册成功！采集了 {owner_detected} 个不同姿势的特征",
                                        parent=parent)
                else:
                    messagebox.showinfo("成功", f"主人面部注册成功！采集了 {owner_detected} 个不同姿势的特征")
                return True
            else:
                if parent:
                    messagebox.showerror("错误", "采集的样本不足，请重试", parent=parent)
                else:
                    messagebox.showerror("错误", "采集的样本不足，请重试")
                return False
        except Exception as e:
            self.logger.error(f"注册主人时出错: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            try:
                if parent:
                    messagebox.showerror("错误", f"注册主人时出错: {e}", parent=parent)
                else:
                    messagebox.showerror("错误", f"注册主人时出错: {e}")
            except:
                pass
            return False
        finally:
            # 确保释放摄像头资源
            try:
                if cap:
                    cap.release()
            except Exception as e:
                self.logger.error(f"释放摄像头资源时出错: {e}")
            # 确保关闭所有OpenCV窗口
            try:
                cv2.destroyAllWindows()
            except Exception as e:
                self.logger.error(f"关闭OpenCV窗口时出错: {e}")

    def set_startup(self, enable):
        try:
            key = winreg.HKEY_CURRENT_USER
            subkey = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(key, subkey, 0, winreg.KEY_SET_VALUE) as reg_key:
                if enable:
                    import sys
                    script_path = os.path.abspath(sys.argv[0])
                    winreg.SetValueEx(reg_key, "PrivacyShield", 0, winreg.REG_SZ, script_path)
                else:
                    try:
                        winreg.DeleteValue(reg_key, "PrivacyShield")
                    except:
                        pass
            return True
        except:
            return False

    def create_tray_icon(self):
        def on_quit(icon, item):
            icon.stop()
            self.stop_protection()
            if self.root:
                self.root.destroy()
            os._exit(0)

        def on_enable():
            self.start_protection()

        def on_disable():
            self.stop_protection()

        def on_config():
            # 确保主窗口存在
            self.create_root()
            # 使用 after 调用，确保在 Tkinter 主线程中执行
            if self.root and self.root.winfo_exists():
                self.root.after(0, self.show_settings)

        def on_wizard():
            # 确保主窗口存在
            self.create_root()
            # 使用 after 调用，确保在 Tkinter 主线程中执行
            if self.root and self.root.winfo_exists():
                self.root.after(0, self.show_config_wizard)

        def on_register():
            # 确保主窗口存在
            self.create_root()
            # 使用 after 调用，确保在 Tkinter 主线程中执行
            if self.root and self.root.winfo_exists():
                self.root.after(0, self.register_owner_wizard)

        image = PILImage.new('RGB', (64, 64), color=(255, 255, 255))
        from PIL import ImageDraw
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill=(0, 128, 255))
        draw.text((20, 20), "防窥", fill=(255, 255, 255))

        menu = (
            item('启用保护', on_enable),
            item('禁用保护', on_disable),
            item('注册主人', on_register),
            item('配置向导', on_wizard),
            item('设置', on_config),
            item('退出', on_quit),
        )

        self.tray_icon = pystray.Icon("防窥屏程序", image, "防窥屏程序", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_settings(self):
        # 确保主窗口存在
        self.create_root()
        if not self.root or not self.root.winfo_exists():
            self.logger.error("主窗口不存在，无法创建设置窗口")
            return

        # 检查设置窗口是否已经存在
        if self.settings_window and self.settings_window.winfo_exists():
            # 窗口已存在，将其置于前台
            self.settings_window.deiconify()
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        # 直接创建窗口，不使用 after，确保在当前线程中执行
        try:
            self.logger.info("正在创建设置窗口")
            settings_window = tk.Toplevel(self.root)
            settings_window.title("防窥屏程序 - 设置")
            settings_window.geometry("600x600")
            settings_window.resizable(False, False)
            settings_window.transient(self.root)
            # 确保窗口可见
            settings_window.deiconify()
            settings_window.lift()
            settings_window.focus_force()

            # 保存窗口引用
            self.settings_window = settings_window

            # 设置窗口关闭回调
            def on_settings_close():
                self.settings_window = None
                settings_window.destroy()

            settings_window.protocol("WM_DELETE_WINDOW", on_settings_close)
            self.logger.info("设置窗口已创建")

            notebook = ttk.Notebook(settings_window)

            basic_frame = ttk.Frame(notebook)
            notebook.add(basic_frame, text="基本设置")

            ttk.Label(basic_frame, text="识别灵敏度：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
            sensitivity_var = tk.DoubleVar(value=self.config['sensitivity'])
            ttk.Scale(basic_frame, from_=0.1, to=1.0, variable=sensitivity_var, orient=tk.HORIZONTAL).grid(row=0,
                                                                                                           column=1,
                                                                                                           padx=10,
                                                                                                           pady=10,
                                                                                                           sticky=tk.W)
            ttk.Label(basic_frame, textvariable=sensitivity_var, width=10).grid(row=0, column=2, padx=10, pady=10)
            ttk.Label(basic_frame, text="灵敏度越高，识别越严格，可能会将主人误判为陌生人",
                      font=("Arial", 10), foreground="gray").grid(row=1, column=0, columnspan=3, padx=10, pady=2,
                                                                  sticky=tk.W)

            ttk.Label(basic_frame, text="响应时间 (秒)：").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
            response_var = tk.DoubleVar(value=self.config['response_time'])
            ttk.Scale(basic_frame, from_=0.1, to=2.0, variable=response_var, orient=tk.HORIZONTAL).grid(row=2, column=1,
                                                                                                        padx=10,
                                                                                                        pady=10,
                                                                                                        sticky=tk.W)
            ttk.Label(basic_frame, textvariable=response_var, width=10).grid(row=2, column=2, padx=10, pady=10)
            ttk.Label(basic_frame, text="响应时间越短，检测到陌生人后黑屏越快",
                      font=("Arial", 10), foreground="gray").grid(row=3, column=0, columnspan=3, padx=10, pady=2,
                                                                  sticky=tk.W)

            startup_var = tk.BooleanVar(value=self.config['startup'])
            ttk.Checkbutton(basic_frame, text="开机自启动", variable=startup_var).grid(row=4, column=0, columnspan=3,
                                                                                       padx=10, pady=10, sticky=tk.W)

            ttk.Label(basic_frame, text="截图保存路径：").grid(row=5, column=0, padx=10, pady=10, sticky=tk.W)
            path_var = tk.StringVar(value=self.config['screenshot_path'])
            ttk.Entry(basic_frame, textvariable=path_var, width=40).grid(row=5, column=1, padx=10, pady=10)

            advanced_frame = ttk.Frame(notebook)
            notebook.add(advanced_frame, text="高级设置")

            ttk.Label(advanced_frame, text="检测区域 (百分比)：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
            ttk.Label(advanced_frame, text="当前设置：全屏", font=("Arial", 10)).grid(row=1, column=0, padx=10, pady=5,
                                                                                     sticky=tk.W)
            ttk.Label(advanced_frame, text="注：检测区域设置功能将在后续版本中提供",
                      font=("Arial", 10), foreground="gray").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)

            def save_settings():
                try:
                    self.config['sensitivity'] = sensitivity_var.get()
                    self.config['response_time'] = response_var.get()
                    self.config['startup'] = startup_var.get()
                    # 保持 show_camera 为 True
                    self.config['show_camera'] = True
                    self.show_camera = True
                    self.config['screenshot_path'] = path_var.get()
                    self.save_config()

                    if self.video_capture:
                        try:
                            self.video_capture.release()
                        except Exception as e:
                            self.logger.error(f"释放摄像头资源时出错: {e}")
                        self.video_capture = None

                    if startup_var.get():
                        self.set_startup(True)
                    else:
                        self.set_startup(False)

                    messagebox.showinfo("成功", "设置已保存")
                    self.settings_window = None
                    settings_window.destroy()
                except Exception as e:
                    self.logger.error(f"保存设置时出错: {e}")
                    messagebox.showerror("错误", f"保存设置时出错: {e}")

            ttk.Button(settings_window, text="保存", command=save_settings).pack(pady=20)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            settings_window.wait_window()
        except Exception as e:
            self.logger.error(f"创建设置窗口时出错: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            try:
                messagebox.showerror("错误", f"创建设置窗口时出错: {e}\n请查看 privacy_shield.log 获取详细信息")
            except:
                pass

    def start_protection(self):
        if self.is_running:
            return

        self.is_running = True

        self.listener = mouse.Listener(on_click=self.on_mouse_click)
        self.listener.start()

        def on_activate():
            if self.is_running:
                self.stop_protection()
            else:
                self.start_protection()

        try:
            self.hotkey_listener = keyboard.GlobalHotKeys({
                '<ctrl>+<alt>+p': on_activate,
            })
            self.hotkey_listener.start()
        except:
            print("热键注册失败，将不使用热键功能")
            self.hotkey_listener = None

        threading.Thread(target=self.detection_loop, daemon=True).start()
        threading.Thread(target=self.check_click_timeout, daemon=True).start()

        if self.tray_icon:
            self.tray_icon.notify("防窥屏程序已启动", "保护已启用")

    def stop_protection(self):
        if not self.is_running:
            return

        self.is_running = False
        self.is_black_screen = False

        if self.video_capture:
            try:
                self.video_capture.release()
            except Exception as e:
                self.logger.error(f"释放摄像头资源时出错: {e}")
            self.video_capture = None

        if self.listener:
            self.listener.stop()

        if self.hotkey_listener:
            try:
                self.hotkey_listener.stop()
            except:
                pass

        if self.black_screen_window:
            self.black_screen_window.destroy()
            self.black_screen_window = None

        if self.tray_icon:
            self.tray_icon.notify("防窥屏程序已停止", "保护已禁用")

    def compare_faces(self, face_hist):
        max_correlation = 0
        for owner_hist in self.owner_encodings:
            try:
                correlation = cv2.compareHist(face_hist, owner_hist, cv2.HISTCMP_CORREL)
                max_correlation = max(max_correlation, correlation)
            except:
                continue
        return max_correlation > self.config['sensitivity']

    def is_valid_face(self, face_roi):
        h, w = face_roi.shape
        aspect_ratio = w / h
        mean_brightness = np.mean(face_roi)
        edges = cv2.Canny(face_roi, 100, 200)
        edge_density = np.sum(edges > 0) / (h * w)

        # 计算面部的纹理特征
        lbp_hist = cv2.calcHist([face_roi], [0], None, [256], [0, 256])
        lbp_std = np.std(lbp_hist)

        if 0.6 < aspect_ratio < 1.2 and 30 < mean_brightness < 220 and edge_density > 0.05 and lbp_std > 10:
            return True
        return False

    def detection_loop(self):
        # 获取人脸检测器
        face_cascade = self.get_face_cascade()
        if not face_cascade:
            self.logger.error("无法加载人脸检测器，检测循环无法启动")
            return

        # 摄像头状态监控
        last_frame_time = time.time()
        frame_timeout = 5  # 5秒内没有收到帧就认为摄像头卡死

        # 窗口状态
        camera_window_open = False

        while self.is_running:
            if not self.video_capture:
                try:
                    # 直接使用DirectShow后端，避免MSMF错误
                    self.video_capture = cv2.VideoCapture(self.camera_index, cv2.CAP_DSHOW)
                    if not self.video_capture.isOpened():
                        # 如果DirectShow失败，尝试默认后端
                        self.video_capture = cv2.VideoCapture(self.camera_index)
                    if self.video_capture.isOpened():
                        # 设置摄像头参数以优化性能
                        self.video_capture.set(cv2.CAP_PROP_FRAME_WIDTH, 320)  # 降低分辨率提高流畅度
                        self.video_capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 240)
                        self.video_capture.set(cv2.CAP_PROP_FPS, 10)  # 降低帧率提高流畅度

                        # 优化摄像头设置，避免闪烁
                        try:
                            # 禁用自动增益，避免闪烁
                            self.video_capture.set(cv2.CAP_PROP_AUTO_GAIN, 0)
                            # 设置固定曝光，避免自动曝光导致的闪烁
                            self.video_capture.set(cv2.CAP_PROP_AUTO_EXPOSURE, 0.25)  # 0.25 表示手动曝光
                            # 设置固定亮度和对比度
                            self.video_capture.set(cv2.CAP_PROP_BRIGHTNESS, 128)
                            self.video_capture.set(cv2.CAP_PROP_CONTRAST, 50)
                            self.video_capture.set(cv2.CAP_PROP_SATURATION, 50)
                        except Exception as e:
                            # 某些摄像头不支持这些设置，忽略错误
                            pass

                        # 验证是否能读取到帧
                        ret, test_frame = self.video_capture.read()
                        if ret and test_frame is not None and test_frame.size > 0:
                            self.logger.info(f"成功打开摄像头 {self.camera_index}")
                        else:
                            self.video_capture.release()
                            self.video_capture = None
                            time.sleep(5)  # 增加重试间隔，减少CPU占用
                            continue
                    else:
                        self.logger.error(f"无法打开摄像头 {self.camera_index}")
                        self.video_capture = None
                        time.sleep(5)  # 增加重试间隔，减少CPU占用
                        continue

                except Exception as e:
                    self.logger.error(f"打开摄像头时出错: {e}")
                    if self.video_capture:
                        try:
                            self.video_capture.release()
                        except:
                            pass
                        self.video_capture = None
                    time.sleep(5)
                    continue

            try:
                # 移除跳帧机制，避免摄像头冲突
                # 直接读取帧

                # 读取帧
                try:
                    ret, frame = self.video_capture.read()
                    current_time = time.time()

                    # 如果帧读取失败，尝试重新初始化摄像头
                    if not ret or frame is None or frame.size == 0:
                        self.logger.warning(f"摄像头 {self.camera_index} 读取帧失败，尝试重新初始化")
                        # 释放摄像头资源，重新初始化
                        try:
                            self.video_capture.release()
                        except Exception as e2:
                            self.logger.error(f"释放摄像头资源时出错: {e2}")
                        self.video_capture = None
                        # 关闭摄像头窗口
                        if camera_window_open:
                            try:
                                cv2.destroyWindow('防窥屏 - 摄像头')
                                camera_window_open = False
                            except:
                                pass
                        time.sleep(2)
                        continue
                except Exception as e:
                    self.logger.error(f"读取摄像头帧时发生异常: {e}")
                    # 释放摄像头资源，重新初始化
                    try:
                        self.video_capture.release()
                    except Exception as e2:
                        self.logger.error(f"释放摄像头资源时出错: {e2}")
                    self.video_capture = None
                    # 关闭摄像头窗口
                    if camera_window_open:
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass
                    time.sleep(2)
                    continue

                # 检查摄像头是否超时
                if current_time - last_frame_time > frame_timeout:
                    self.logger.warning(f"摄像头 {self.camera_index} 超时，可能已卡死")
                    # 释放摄像头资源，重新初始化
                    try:
                        self.video_capture.release()
                    except Exception as e:
                        self.logger.error(f"释放摄像头资源时出错: {e}")
                    self.video_capture = None
                    # 关闭摄像头窗口
                    if camera_window_open:
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass
                    time.sleep(2)
                    continue

                if not ret or frame is None or frame.size == 0:
                    # 帧读取失败，可能是摄像头断开或出错
                    self.logger.warning(f"摄像头 {self.camera_index} 读取帧失败")
                    # 释放摄像头资源，稍后重新打开
                    try:
                        self.video_capture.release()
                    except Exception as e:
                        self.logger.error(f"释放摄像头资源时出错: {e}")
                    self.video_capture = None
                    # 关闭摄像头窗口
                    if camera_window_open:
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass
                    time.sleep(2)
                    continue

                # 更新最后帧时间
                last_frame_time = current_time

                # 检查帧是否是全黑的
                try:
                    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                    mean_brightness = np.mean(gray)
                    if mean_brightness < 10:  # 全黑帧
                        # 跳过全黑帧，继续下一帧
                        self.logger.debug(f"检测到全黑帧，亮度: {mean_brightness}")
                        continue
                except Exception as e:
                    self.logger.error(f"处理帧时出错: {e}")
                    # 继续下一帧
                    continue

                if self.is_black_screen:
                    continue

                # 预处理图像以提高检测效果
                try:
                    gray = cv2.equalizeHist(gray)  # 直方图均衡化，提高对比度
                except Exception as e:
                    self.logger.error(f"直方图均衡化时出错: {e}")
                    # 继续使用原始灰度图像
                    pass

                # 计算环境光线亮度
                brightness = mean_brightness  # 复用之前计算的亮度值
                # 根据亮度自动调整灵敏度
                adaptive_sensitivity = self.adjust_sensitivity(brightness)

                # 改进的人脸检测参数
                faces = []
                try:
                    # 优化人脸检测参数，提高性能
                    faces = face_cascade.detectMultiScale(
                        gray,
                        scaleFactor=1.3,  # 增大缩放因子，减少计算量
                        minNeighbors=3,  # 减少邻居数，提高检测速度
                        minSize=(80, 80)  # 减小最小人脸尺寸
                    )
                except Exception as e:
                    # 如果检测器出错，重新加载检测器
                    self.logger.warning(f"人脸检测器出错，尝试重新加载: {e}")
                    self.face_cascade = None
                    time.sleep(0.5)
                    face_cascade = self.get_face_cascade()
                    if not face_cascade:
                        self.logger.error("重新加载人脸检测器失败")
                        time.sleep(1)
                        continue

                valid_faces = []
                for (x, y, w, h) in faces:
                    # 调整人脸大小阈值，与检测参数一致
                    if w < 80 or h < 80:
                        continue

                    face_roi = gray[y:y + h, x:x + w]

                    # 简化验证逻辑，提高检测率
                    if w > 60 and h > 60:  # 只要人脸大小足够就认为有效
                        valid_faces.append((x, y, w, h))

                self.owner_detected = False
                if len(valid_faces) > 0:
                    for (x, y, w, h) in valid_faces:
                        face_roi = gray[y:y + h, x:x + w]
                        try:
                            face_roi = cv2.resize(face_roi, (100, 100))

                            hist = cv2.calcHist([face_roi], [0], None, [256], [0, 256])
                            hist = cv2.normalize(hist, hist).flatten()

                            if self.compare_faces(hist, adaptive_sensitivity):
                                self.owner_detected = True
                                break
                        except Exception as e:
                            self.logger.error(f"处理人脸特征时出错: {e}")
                            continue

                # 绘制检测结果
                try:
                    if self.owner_registered and not self.owner_detected and len(valid_faces) > 0:
                        for (x, y, w, h) in valid_faces:
                            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                        cv2.putText(frame, "陌生人！", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

                        if self.stranger_start_time is None:
                            self.stranger_start_time = time.time()
                        elif time.time() - self.stranger_start_time >= self.config['response_time']:
                            timestamp = time.strftime("%Y%m%d_%H%M%S")
                            screenshot_path = os.path.join(self.config['screenshot_path'], f"stranger_{timestamp}.jpg")
                            cv2.imwrite(screenshot_path, frame, [cv2.IMWRITE_JPEG_QUALITY, 95])
                            print(f"已保存陌生人截图: {screenshot_path}")

                            self.show_black_screen()
                            self.stranger_start_time = None
                    else:
                        if self.owner_detected:
                            cv2.putText(frame, "主人", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
                        # 显示当前灵敏度
                        if self.show_camera:
                            cv2.putText(frame, f"灵敏度: {adaptive_sensitivity:.2f}", (10, 60),
                                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 0), 2)
                        self.stranger_start_time = None
                except Exception as e:
                    self.logger.error(f"绘制检测结果时出错: {e}")

                # 显示摄像头窗口
                if self.show_camera:
                    try:
                        # 检查窗口是否存在
                        if not camera_window_open:
                            cv2.namedWindow('防窥屏 - 摄像头')
                            camera_window_open = True

                        cv2.imshow('防窥屏 - 摄像头', frame)
                        # 使用非阻塞的方式等待按键
                        key = cv2.waitKey(1) & 0xFF
                        if key == ord('q'):
                            self.show_camera = False
                            # 关闭窗口
                            try:
                                cv2.destroyWindow('防窥屏 - 摄像头')
                                camera_window_open = False
                            except:
                                pass
                        # 检查窗口是否被用户关闭
                        if camera_window_open:
                            try:
                                if cv2.getWindowProperty('防窥屏 - 摄像头', cv2.WND_PROP_VISIBLE) < 1:
                                    camera_window_open = False
                                    self.show_camera = False
                            except:
                                camera_window_open = False
                                self.show_camera = False
                    except Exception as e:
                        self.logger.error(f"显示摄像头窗口时出错: {e}")
                        # 关闭窗口
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass
                        self.show_camera = False
                else:
                    # 如果不需要显示摄像头，确保窗口已关闭
                    if camera_window_open:
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass

            except Exception as e:
                self.logger.error(f"检测循环错误: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
                # 检查是否是摄像头相关错误
                if "videoio" in str(e) or "cap_msmf" in str(e) or "grabFrame" in str(e):
                    self.logger.warning("检测到摄像头错误，尝试重新初始化摄像头")
                    if self.video_capture:
                        self.video_capture.release()
                        self.video_capture = None
                    # 关闭摄像头窗口
                    if camera_window_open:
                        try:
                            cv2.destroyWindow('防窥屏 - 摄像头')
                            camera_window_open = False
                        except:
                            pass
                    time.sleep(3)
                else:
                    time.sleep(0.1)
                continue

            time.sleep(0.05)  # 稍微减少等待时间

    def adjust_sensitivity(self, brightness):
        """根据环境光线亮度自动调整灵敏度"""
        # 亮度范围通常在0-255之间
        if brightness < 50:  # 光线较暗
            # 降低灵敏度，使识别更宽松
            return max(0.3, self.config['sensitivity'] - 0.15)
        elif brightness > 200:  # 光线较亮
            # 提高灵敏度，使识别更严格
            return min(0.8, self.config['sensitivity'] + 0.15)
        else:  # 光线适中
            return self.config['sensitivity']

    def compare_faces(self, face_hist, sensitivity=None):
        if sensitivity is None:
            sensitivity = self.config['sensitivity']

        max_correlation = 0
        for owner_hist in self.owner_encodings:
            try:
                correlation = cv2.compareHist(face_hist, owner_hist, cv2.HISTCMP_CORREL)
                max_correlation = max(max_correlation, correlation)
            except:
                continue
        return max_correlation > sensitivity

    def show_black_screen(self):
        self.create_root()

        self.is_black_screen = True
        self.click_count = 0

        self.black_screen_window = tk.Toplevel(self.root)
        self.black_screen_window.attributes('-fullscreen', True)
        self.black_screen_window.configure(bg='black')
        self.black_screen_window.attributes('-topmost', True)
        self.black_screen_window.overrideredirect(True)

        label = tk.Label(self.black_screen_window, text='检测到陌生人！\\n点击鼠标左键三下解除',
                         fg='white', bg='black', font=('Arial', 24))
        label.pack(expand=True)

    def on_mouse_click(self, x, y, button, pressed):
        if not self.is_black_screen:
            return

        if button == mouse.Button.left and pressed:
            current_time = time.time()

            if current_time - self.last_click_time > 2:
                self.click_count = 0

            self.click_count += 1
            self.last_click_time = current_time

            if self.click_count >= 3:
                self.hide_black_screen()

    def check_click_timeout(self):
        while self.is_running:
            if self.is_black_screen and time.time() - self.last_click_time > 2:
                self.click_count = 0
            time.sleep(0.5)

    def hide_black_screen(self):
        self.is_black_screen = False
        self.click_count = 0

        if self.black_screen_window:
            self.black_screen_window.destroy()
            self.black_screen_window = None

    def run(self):
        try:
            # 创建主窗口
            self.create_root()

            def initialize():
                try:
                    self.start_protection()
                    self.create_tray_icon()
                    if self.need_config_wizard:
                        self.show_config_wizard()
                except Exception as e:
                    self.logger.error(f"初始化过程出错: {e}")
                    import traceback
                    self.logger.error(traceback.format_exc())
                    # 尝试重新初始化
                    self.root.after(5000, initialize)

            self.root.after(100, initialize)

            try:
                self.root.mainloop()
            except KeyboardInterrupt:
                self.stop_protection()
            except Exception as e:
                self.logger.error(f"主循环出错: {e}")
                import traceback
                self.logger.error(traceback.format_exc())
                # 尝试重启程序
                self.stop_protection()
                self.run()
        except Exception as e:
            self.logger.critical(f"程序启动失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            # 显示错误信息
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            tk.messagebox.showerror("错误", f"程序启动失败: {e}\n请查看 privacy_shield.log 获取详细信息")
            root.destroy()


if __name__ == '__main__':
    app = PrivacyShield()
    app.run()

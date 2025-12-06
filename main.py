import os
import time
from winotify import Notification
import requests
import subprocess
import sys
import yaml

def get_base_dir():
    # 打包后使用 sys.executable，开发运行时使用 __file__
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def load_config(filename=None):
    """
    读取 YAML 配置并返回 dict。
    若文件不存在或解析失败，则在程序目录写入并返回默认配置（等价于参上 config.yaml）。
    filename 为 None 时在脚本目录查找 config.yaml。
    """
    base = get_base_dir()
    if filename is None:
        filename = os.path.join(base, "config.yaml")

    default = {
        "Use_excternal_path:":"No",
        "Path": "",
        "lauch_when_device_start": "No",
        "check_interval": 1000,
        "max_check_times": 3,
        "use_notification": "Yes"
    }

    try:
        with open(filename, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
            if not isinstance(data, dict):
                raise ValueError("config is not a mapping")
            return data
    except Exception:
        # 文件不存在或解析失败：尝试写入默认配置并返回默认 dict
        try:
            with open(filename, "w", encoding="utf-8") as f:
                yaml.safe_dump(default, f, allow_unicode=True, default_flow_style=False)
        except Exception:
            pass
        return default

def get_config_var(cfg, part, default=None):
    """
    如果不存在返回 default
    """
    if not cfg:
        return default
    if part not in cfg:
        return default
    return cfg[part]

def schtask_exists(task_name="TimeChecker"):
    """检查计划任务是否已存在"""
    try:
        res = subprocess.run(["schtasks", "/Query", "/TN", task_name],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return res.returncode == 0
    except Exception:
        return False

def create_schtask(task_name="TimeChecker"):
    """
    创建登录时启动的计划任务，/RL HIGHEST 可设置以最高权限运行（需管理员）。
    幂等性由调用方保证（会先检查是否存在）。
    """
    exe = os.path.abspath(sys.executable)
    # /TR 传入要执行的可执行路径（带引号）
    cmd = [
        "schtasks", "/Create", "/SC", "ONLOGON", "/TN", task_name,
        "/TR", f'"{exe}"', "/RL", "HIGHEST", "/F"
    ]
    try:
        subprocess.run(" ".join(cmd), check=True, shell=True)
        return True, "Created"
    except subprocess.CalledProcessError as e:
        return False, str(e)
    except Exception as e:
        return False, str(e)

def delete_schtask(task_name="TimeChecker"):
    cmd = ["schtasks", "/Delete", "/TN", task_name, "/F"]
    try:
        subprocess.run(cmd, check=True, shell=False)
        return True, "Deleted"
    except subprocess.CalledProcessError as e:
        return False, str(e)
    except Exception as e:
        return False, str(e)

def ensure_schtask_installed(task_name="TimeChecker"):
    """如果不存在则创建计划任务（返回 (ok, msg)）。"""
    if schtask_exists(task_name):
        return True, "exists"
    return create_schtask(task_name)

def ensure_schtask_removed(task_name="TimeChecker"):
    """如果存在则删除计划任务（返回 (ok, msg)）。"""
    if not schtask_exists(task_name):
        return True, "not found"
    return delete_schtask(task_name)

def is_connected():
    try:
        requests.get("https://www.baidu.com", timeout=5)
        return True
    except requests.RequestException:
        return False

def send_notification(title, message):
    enble = get_config_var(load_config(), "use_notification", "Yes")
    if str(enble).lower() in ("yes", "true", "1"):
        toast = Notification(app_id="Time Checker",
                            title=title,
                            msg=message,
                            duration="short")
        toast.show()
    else:
        return

def main():
    config = load_config()
    # 自启动设置（仅使用 schtasks，不再生成 startup bat）
    auto = get_config_var(config, "lauch_when_device_start", "No")
    if str(auto).lower() in ("yes", "true", "1"):
        ok, msg = ensure_schtask_installed()
        if not ok:
            send_notification("Autostart Failed", f"Failed to create scheduled task: {msg}")
    else:
        ok, msg = ensure_schtask_removed()
        if not ok:
            send_notification("Autostart Cleanup Failed", f"Failed to remove scheduled task: {msg}")

    max_check_times = int(get_config_var(config, "max_check_times", 3))
    check_interval = int(get_config_var(config, "check_interval", 1000))
    for i in range(max_check_times):
        if is_connected():
            try:
                if get_config_var(config, "Use_excternal_path", "No").lower() in ("yes", "true", "1"):
                    path = get_config_var(config, "Path", "")
                    if not path:
                        send_notification("Failure", "No valid path configured for time update executable.")
                        return
                    os.startfile(path)
                else:
                    os.system("w32tm /resync")
                send_notification("Success", "The time has been checked successfully.")
            except Exception as e:
                send_notification("Failure", f"Failed to launch Time checker: {e}")
            break
        else:
            if i < max_check_times - 1:
                send_notification("Failure", f"Network connection failed. Remain attempts: {max_check_times - i - 1}")
                time.sleep(check_interval)
            else:
                send_notification("Failure", "Time check failed.")

if __name__ == "__main__":
    main()
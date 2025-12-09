import os
import time
from winotify import Notification
import requests
import subprocess
import sys
import yaml
import tempfile
from datetime import datetime, timezone
import xml.etree.ElementTree as ET

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

def _build_task_xml(task_name: str, command: str, arguments: str = "") -> bytes:
    """
    生成 Task XML（bytes，UTF-16 编码），UserId 使用当前用户，确保在电池上也能启动。
    返回 UTF-16 编码的 XML 内容。
    """
    ns = "http://schemas.microsoft.com/windows/2004/02/mit/task"
    ET.register_namespace("", ns)
    T = lambda tag: f"{{{ns}}}{tag}"

    task = ET.Element(T("Task"), {"version": "1.2"})
    reg = ET.SubElement(task, T("RegistrationInfo"))
    ET.SubElement(reg, T("Date")).text = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    ET.SubElement(reg, T("Author")).text = os.environ.get("USERNAME", "Unknown")

    # 触发器：仅使用 LogonTrigger（用户登录时触发），删除其他触发器
    triggers = ET.SubElement(task, T("Triggers"))
    logon = ET.SubElement(triggers, T("LogonTrigger"))
    ET.SubElement(logon, T("Enabled")).text = "true"
    # 可选延迟，避免登录瞬间环境未就绪
    ET.SubElement(logon, T("Delay")).text = "PT10S"

    # Principal 使用当前用户
    principals = ET.SubElement(task, T("Principals"))
    principal = ET.SubElement(principals, T("Principal"), {"id": "Author"})
    domain = os.environ.get("USERDOMAIN", "")
    user = os.environ.get("USERNAME", "")
    userid = f"{domain}\\{user}" if domain else user
    ET.SubElement(principal, T("UserId")).text = userid
    ET.SubElement(principal, T("LogonType")).text = "InteractiveToken"
    # 要求以最高权限运行
    ET.SubElement(principal, T("RunLevel")).text = "HighestAvailable"

    # Settings: 允许在电池上运行（DisallowStartIfOnBatteries=false），并尽量保守设置
    settings = ET.SubElement(task, T("Settings"))
    idle = ET.SubElement(settings, T("IdleSettings"))
    ET.SubElement(idle, T("Duration")).text = "PT10M"
    ET.SubElement(idle, T("WaitTimeout")).text = "PT1H"
    ET.SubElement(idle, T("StopOnIdleEnd")).text = "true"
    ET.SubElement(idle, T("RestartOnIdle")).text = "false"

    ET.SubElement(settings, T("MultipleInstancesPolicy")).text = "IgnoreNew"
    ET.SubElement(settings, T("DisallowStartIfOnBatteries")).text = "false"
    ET.SubElement(settings, T("StopIfGoingOnBatteries")).text = "false"
    ET.SubElement(settings, T("AllowHardTerminate")).text = "true"
    ET.SubElement(settings, T("StartWhenAvailable")).text = "false"
    ET.SubElement(settings, T("RunOnlyIfNetworkAvailable")).text = "false"
    ET.SubElement(settings, T("AllowStartOnDemand")).text = "true"
    ET.SubElement(settings, T("Enabled")).text = "true"
    ET.SubElement(settings, T("Hidden")).text = "false"
    ET.SubElement(settings, T("RunOnlyIfIdle")).text = "false"
    ET.SubElement(settings, T("WakeToRun")).text = "false"
    ET.SubElement(settings, T("ExecutionTimeLimit")).text = "P3D"
    ET.SubElement(settings, T("Priority")).text = "7"

    # Actions：执行当前 Python 可执行文件（或直接执行脚本/可执行）
    actions = ET.SubElement(task, T("Actions"), {"Context": "Author"})
    exec_el = ET.SubElement(actions, T("Exec"))
    ET.SubElement(exec_el, T("Command")).text = command
    if arguments:
        ET.SubElement(exec_el, T("Arguments")).text = arguments

    tree = ET.ElementTree(task)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="wb") as tf:
        tree.write(tf, encoding="utf-16", xml_declaration=True)
        tf.flush()
        tfpath = tf.name
    try:
        with open(tfpath, "rb") as f:
            content = f.read()
    finally:
        try:
            os.unlink(tfpath)
        except Exception:
            pass
    return content

def create_schtask(task_name="TimeChecker"):
    """
    使用动态生成的 XML 创建计划任务（保证 UserId 与路径匹配，并允许在电池上运行）。
    如果生成 XML 或创建失败，则回退到 /SC ONLOGON /TR 方式。
    """
    try:
        # 使用当前 Python 可执行作为命令；如果你想运行脚本可替换为脚本路径并传参
        command = os.path.abspath(sys.executable)
        # 如果是直接运行脚本，可将 arguments 设置为脚本路径：
        script_path = os.path.join(get_base_dir(), os.path.basename(sys.argv[0]))
        arguments = f'"{script_path}"' if os.path.exists(script_path) else ""
        xml_bytes = _build_task_xml(task_name, command, arguments)

        # 写临时 XML 文件并用 schtasks /Create /XML 创建
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="wb") as tf:
            tf.write(xml_bytes)
            tf.flush()
            xml_file = tf.name
        try:
            cmd = ["schtasks", "/Create", "/TN", task_name, "/XML", xml_file, "/F"]
            subprocess.run(cmd, check=True, shell=False)
            return True, "Created from generated XML"
        finally:
            try:
                os.unlink(xml_file)
            except Exception:
                pass
    except subprocess.CalledProcessError as e:
        # 回退到 /TR 创建（保留行为）
        try:
            exe = os.path.abspath(sys.executable)
            cmd = ["schtasks", "/Create", "/SC", "ONLOGON", "/TN", task_name,
                   "/TR", f'"{exe}"', "/RL", "HIGHEST", "/F"]
            subprocess.run(" ".join(cmd), check=True, shell=True)
            return True, f"Created by /TR fallback (xml error: {e})"
        except Exception as e2:
            return False, f"both methods failed: {e} / {e2}"
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
                    send_notification("Success", "The time has been checked successfully.")
                else:
                    # 简化：尝试启动 w32time 服务（若尚未运行），然后执行 w32tm /resync
                    try:
                        q = subprocess.run(["sc", "query", "w32time"], capture_output=True, text=True)
                        out = (q.stdout or "") + (q.stderr or "")
                        if "RUNNING" not in out:
                            subprocess.run(["sc", "start", "w32time"], capture_output=True, text=True)
                            # 等待短暂确认
                            for _ in range(5):
                                time.sleep(1)
                                q = subprocess.run(["sc", "query", "w32time"], capture_output=True, text=True)
                                if "RUNNING" in (q.stdout or ""):
                                    break
                        proc = subprocess.run(["w32tm", "/resync"], capture_output=True, text=True)
                        if proc.returncode != 0:
                            raise RuntimeError(f"w32tm failed: {proc.returncode} {proc.stdout} {proc.stderr}")
                        send_notification("Success", "The time has been checked successfully.")
                    except Exception as e:
                        send_notification("Failure", f"w32tm failed or service not running: {e}")
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
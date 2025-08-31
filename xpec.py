import platform
import subprocess
import psutil
import webbrowser
from datetime import datetime
import os
import re
import json
import sys

# Optional image generation deps
try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Windows-specific imports
if platform.system() == "Windows":
    try:
        import comtypes
        from comtypes.client import CreateObject
    except ImportError:
        comtypes = None
    import wmi
    import winreg

def _is_debug() -> bool:
    # Enable via env: XPEC_DEBUG_MOBO=1 or XPEC_DEBUG=1
    v = os.getenv("XPEC_DEBUG_MOBO") or os.getenv("XPEC_DEBUG")
    return str(v).lower() in {"1", "true", "yes", "on"}

def _debug_print(msg: str):
    if _is_debug():
        try:
            print(msg)
        except Exception:
            pass

# GPU debug controls
GPU_DEBUG = []  # collected lines for HTML
def _is_debug_gpu() -> bool:
    v = os.getenv("XPEC_DEBUG_GPU") or os.getenv("XPEC_DEBUG")
    return str(v).lower() in {"1", "true", "yes", "on"}

def _debug_print_gpu(msg: str):
    if _is_debug_gpu():
        try:
            print(msg)
        except Exception:
            pass

def get_system_info():
    system_info = {}
    if platform.system() == "Windows":
        # Prefer a human-friendly motherboard name
        mobo_name = None
        debug_source = None
        debug_reg_manu = debug_reg_product = debug_reg_sysmanu = debug_reg_sysprod = debug_product_pretty = debug_vendor_short = None
        debug_wmi_manufacturer = debug_wmi_product = debug_wmi_version = None
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\\DESCRIPTION\\System\\BIOS") as key:
                manu = (winreg.QueryValueEx(key, "BaseBoardManufacturer")[0] or "").strip()
                product = (winreg.QueryValueEx(key, "BaseBoardProduct")[0] or "").strip()
                sysprod = None
                try:
                    sysprod = (winreg.QueryValueEx(key, "SystemProductName")[0] or "").strip()
                except Exception:
                    sysprod = None
                try:
                    sysmanu = (winreg.QueryValueEx(key, "SystemManufacturer")[0] or "").strip()
                except Exception:
                    sysmanu = None
                # Prefer BaseBoardProduct; fallback to SystemProductName only if BaseBoardProduct is empty or generic
                bad_tokens = {"TO BE FILLED BY O.E.M.", "SYSTEM PRODUCT NAME", "DEFAULT STRING", "NOT SPECIFIED", "UNKNOWN", "N/A"}
                product_pretty = product if product and product.upper() not in bad_tokens else (sysprod or "")
                # Clean patterns like " (MS-7E49)"
                if product_pretty:
                    product_pretty = re.sub(r"\s*\(MS-[0-9A-F]+\)$", "", product_pretty, flags=re.I)
                if manu or product:
                    vendor = _short_vendor(manu)
                    mobo_name = f"{vendor} {product_pretty}".strip()
                    debug_source = "registry"
                    debug_reg_manu = manu
                    debug_reg_product = product
                    debug_reg_sysmanu = sysmanu
                    debug_reg_sysprod = sysprod
                    debug_product_pretty = product_pretty
                    debug_vendor_short = vendor
        except Exception:
            pass
        if not mobo_name:
            try:
                c = wmi.WMI()
                bb = c.Win32_BaseBoard()[0]
                manu = (bb.Manufacturer or "").strip()
                product = (bb.Product or "").strip()
                try:
                    debug_wmi_version = (bb.Version or "").strip()
                except Exception:
                    debug_wmi_version = None
                if manu or product:
                    vendor = _short_vendor(manu)
                    mobo_name = f"{vendor} {product}".strip()
                    debug_source = "wmi"
                    debug_wmi_manufacturer = manu
                    debug_wmi_product = product
            except Exception:
                pass
        system_info["Motherboard"] = mobo_name or "N/A"
        if _is_debug():
            # Attach debug fields for HTML and print to console
            system_info["Debug: Source"] = debug_source or "N/A"
            system_info["Debug: Reg BaseBoardManufacturer"] = debug_reg_manu or "N/A"
            system_info["Debug: Reg BaseBoardProduct"] = debug_reg_product or "N/A"
            system_info["Debug: Reg SystemManufacturer"] = debug_reg_sysmanu or "N/A"
            system_info["Debug: Reg SystemProductName"] = debug_reg_sysprod or "N/A"
            system_info["Debug: Reg Product Pretty"] = debug_product_pretty or "N/A"
            system_info["Debug: Vendor Short"] = debug_vendor_short or "N/A"
            system_info["Debug: WMI BaseBoard.Manufacturer"] = debug_wmi_manufacturer or "N/A"
            system_info["Debug: WMI BaseBoard.Product"] = debug_wmi_product or "N/A"
            system_info["Debug: WMI BaseBoard.Version"] = debug_wmi_version or "N/A"
            _debug_print(f"[MOBO DEBUG] source={debug_source} | reg=({debug_reg_manu} {debug_reg_product}) sys=({debug_reg_sysmanu} {debug_reg_sysprod}) pretty={debug_product_pretty} | wmi=({debug_wmi_manufacturer} {debug_wmi_product} v{debug_wmi_version}) | chosen='{system_info['Motherboard']}'")
    else:
        try:
            with open("/sys/devices/virtual/dmi/id/board_vendor", "r") as f:
                manu = f.read().strip()
            with open("/sys/devices/virtual/dmi/id/board_name", "r") as f:
                product = f.read().strip()
            system_info["Motherboard"] = f"{manu} {product}".strip()
            if _is_debug():
                system_info["Debug: Linux board_vendor"] = manu or "N/A"
                system_info["Debug: Linux board_name"] = product or "N/A"
                _debug_print(f"[MOBO DEBUG] linux board_vendor='{manu}' board_name='{product}' chosen='{system_info['Motherboard']}'")
        except FileNotFoundError:
            system_info["Motherboard"] = "N/A"

    system_info["OS"] = f"{platform.system()} {platform.release()}"
    return system_info

def get_cpu_info():
    cpu_info = {}
    if platform.system() == "Windows":
        try:
            c = wmi.WMI()
            cpu = c.Win32_Processor()[0]
            cpu_info["Model"] = cpu.Name
            cpu_info["Cores"] = cpu.NumberOfCores
            cpu_info["Threads"] = cpu.NumberOfLogicalProcessors
            try:
                mhz = float(cpu.MaxClockSpeed)
                cpu_info["Max Clock"] = f"{mhz/1000:.2f} GHz"
            except Exception:
                pass
        except Exception:
            pass
    else:
        try:
            with open("/proc/cpuinfo", "r") as f:
                for line in f:
                    if "model name" in line:
                        cpu_info["Model"] = line.split(":")[1].strip()
                        break
            cpu_info["Cores"] = psutil.cpu_count(logical=False)
            cpu_info["Threads"] = psutil.cpu_count(logical=True)
            freq = psutil.cpu_freq()
            if freq and freq.max:
                cpu_info["Max Clock"] = f"{freq.max/1000:.2f} GHz"
        except Exception:
            pass

    if not cpu_info.get("Model"):
        cpu_info["Model"] = platform.processor() or "N/A"
    
    return cpu_info

def get_ram_info():
    ram_info = {"Total RAM": f"{psutil.virtual_memory().total / (1024**3):.1f} GB", "Modules": []}
    
    if platform.system() == "Windows":
        try:
            c = wmi.WMI()
            modules = c.Win32_PhysicalMemory()
            for i, module in enumerate(modules, 1):
                manu = (module.Manufacturer or "").strip()
                if not manu or manu.upper() in {"UNKNOWN", "N/A", "UNDEFINED", "NOT SPECIFIED", "INVALID"}:
                    manu = "N/A"
                # Prefer ConfiguredClockSpeed, fallback to Speed
                speed_val = None
                try:
                    speed_val = int(module.ConfiguredClockSpeed) if module.ConfiguredClockSpeed else None
                except Exception:
                    speed_val = None
                if not speed_val:
                    try:
                        speed_val = int(module.Speed) if module.Speed else None
                    except Exception:
                        speed_val = None
                speed_str = f"{speed_val} MHz" if speed_val else "N/A"
                part = (module.PartNumber or "").strip() or "N/A"
                cap_gb = "N/A"
                try:
                    cap_gb = f"{int(module.Capacity) / (1024**3):.1f} GB"
                except Exception:
                    pass
                ram_info["Modules"].append({
                    "Module": i,
                    "Manufacturer": manu,
                    "Capacity": cap_gb,
                    "Speed": speed_str,
                    "Part Number": part
                })
        except Exception:
            pass
    else:
        try:
            result = subprocess.check_output("sudo dmidecode --type 17", shell=True, text=True, stderr=subprocess.DEVNULL)
            modules = []
            current_module = {}
            for line in result.split("\n"):
                if "Size:" in line and "No Module Installed" not in line:
                    current_module["Capacity"] = line.split(":")[1].strip()
                elif "Speed:" in line:
                    current_module["Speed"] = line.split(":")[1].strip()
                elif "Part Number:" in line:
                    current_module["Part Number"] = line.split(":")[1].strip()
                elif "Manufacturer:" in line:
                    current_module["Manufacturer"] = line.split(":")[1].strip()
                elif "Memory Device" in line and current_module:
                    modules.append(current_module)
                    current_module = {}
            
            for i, module in enumerate(modules, 1):
                ram_info["Modules"].append({
                    "Module": i,
                    "Manufacturer": module.get("Manufacturer", "N/A"),
                    "Capacity": module.get("Capacity", "N/A"),
                    "Speed": module.get("Speed", "N/A"),
                    "Part Number": module.get("Part Number", "N/A")
                })
        except Exception:
            pass

    return ram_info

def get_gpu_info():
    gpu_info = []
    global GPU_DEBUG
    GPU_DEBUG = []
    
    if platform.system() == "Windows":
        if _is_debug_gpu():
            _debug_print_gpu("[GPU DEBUG] start detection (Windows)")
            GPU_DEBUG.append(f"Env: comtypes={'yes' if comtypes else 'no'}; NVML=optional; WMI=yes")
        # DXGI for dedicated video memory
        dxgi_gpus = []
        if comtypes:
            try:
                dxgi = CreateObject("DXGI.DXGIFactory")
                idx = 0
                while True:
                    try:
                        adapter = dxgi.EnumAdapters(idx)
                    except Exception:
                        break
                    try:
                        desc = adapter.GetDesc()
                        try:
                            mem_bytes = int(desc.DedicatedVideoMemory)
                        except Exception:
                            mem_bytes = 0
                        vram_str = _fmt_gb_from_bytes(mem_bytes)
                        model = desc.Description.strip()
                        dxgi_gpus.append({
                            "Model": model,
                            "VRAM": vram_str,
                            "_bytes": mem_bytes,
                        })
                        if _is_debug_gpu():
                            GPU_DEBUG.append(f"DXGI[{idx}]: model='{model}' mem_bytes={mem_bytes} -> {vram_str}")
                    except Exception:
                        pass
                    idx += 1
            except Exception as e:
                if _is_debug_gpu():
                    GPU_DEBUG.append(f"DXGI: failed to create factory: {e}")
        
        # NVML (NVIDIA) for accurate VRAM
        nvml_gpus = []
        try:
            import pynvml as N
            N.nvmlInit()
            try:
                count = N.nvmlDeviceGetCount()
                if _is_debug_gpu():
                    GPU_DEBUG.append(f"NVML: init ok, count={count}")
                for i in range(count):
                    h = N.nvmlDeviceGetHandleByIndex(i)
                    raw_name = N.nvmlDeviceGetName(h)
                    name = raw_name.decode() if hasattr(raw_name, "decode") else str(raw_name)
                    mem = int(N.nvmlDeviceGetMemoryInfo(h).total)
                    vram_str = _fmt_gb_from_bytes(mem)
                    nvml_gpus.append({
                        "Model": name,
                        "VRAM": vram_str,
                        "_bytes": mem,
                    })
                    if _is_debug_gpu():
                        GPU_DEBUG.append(f"NVML[{i}]: model='{name}' mem_bytes={mem} -> {vram_str}")
            finally:
                try:
                    N.nvmlShutdown()
                except Exception:
                    pass
        except Exception as e:
            if _is_debug_gpu():
                GPU_DEBUG.append(f"NVML: not available: {e}")

        # WMI to enumerate GPUs and merge info
        try:
            c = wmi.WMI()
            wmi_list = []
            for gpu in c.Win32_VideoController():
                name = gpu.Name
                if not name or "Microsoft" in name:
                    continue
                vram = None
                chosen = None
                # Prefer NVML match
                match_nvml = next((g for g in nvml_gpus if g["Model"] in name or name in g["Model"]), None)
                if match_nvml:
                    vram = match_nvml["VRAM"]
                    chosen = "nvml"
                # Fallback to DXGI match
                if not vram:
                    match_dxgi = next((g for g in dxgi_gpus if g["Model"] in name or name in g["Model"]), None)
                    if match_dxgi:
                        vram = match_dxgi["VRAM"]
                        chosen = "dxgi"
                # Fallback to AdapterRAM
                if not vram and getattr(gpu, "AdapterRAM", None):
                    try:
                        mem = int(gpu.AdapterRAM)
                        vram = _fmt_gb_from_bytes(mem)
                        chosen = "wmi"
                    except Exception:
                        vram = None
                wmi_list.append({"Model": name, "VRAM": vram or "N/A"})
                if _is_debug_gpu():
                    aram = None
                    try:
                        aram = int(gpu.AdapterRAM)
                    except Exception:
                        aram = None
                    GPU_DEBUG.append(
                        f"WMI: name='{name}' AdapterRAM={aram} match_nvml={'yes' if match_nvml else 'no'} match_dxgi={'yes' if 'match_dxgi' in locals() and match_dxgi else 'no'} chosen={chosen or 'none'} -> {vram or 'N/A'}"
                    )

            if wmi_list:
                gpu_info = wmi_list
            elif nvml_gpus:
                gpu_info = [{"Model": g["Model"], "VRAM": g["VRAM"]} for g in nvml_gpus]
            else:
                gpu_info = [{"Model": g["Model"], "VRAM": g["VRAM"]} for g in dxgi_gpus]
        except Exception as e:
            if nvml_gpus:
                gpu_info = [{"Model": g["Model"], "VRAM": g["VRAM"]} for g in nvml_gpus]
            else:
                gpu_info = [{"Model": g["Model"], "VRAM": g["VRAM"]} for g in dxgi_gpus]
            if _is_debug_gpu():
                GPU_DEBUG.append(f"WMI: failed to enumerate: {e}")
    else:
        # Linux implementation
        try:
            result = subprocess.check_output("lspci | grep -i 'vga|3d|display'", shell=True, text=True)
            for line in result.split("\n"):
                if line.strip():
                    model = line.split(":")[2].strip()
                    gpu_info.append({"Model": model, "VRAM": "N/A"})
        except Exception:
            pass

    if _is_debug_gpu():
        _debug_print_gpu(f"[GPU DEBUG] final gpu_info={gpu_info}")
        if not gpu_info:
            GPU_DEBUG.append("Final: no GPUs detected")
        else:
            for i, g in enumerate(gpu_info):
                GPU_DEBUG.append(f"Final[{i}]: model='{g.get('Model','')}' vram='{g.get('VRAM','')}'")

    return gpu_info

def get_disk_info():
    disk_info = []
    if platform.system() == "Windows":
        # 1) Preferred: MSFT_PhysicalDisk (CIM) -> MediaType
        try:
            c2 = wmi.WMI(namespace=r"root\\Microsoft\\Windows\\Storage")
            for pd in c2.MSFT_PhysicalDisk():
                # Size
                try:
                    size_val = int(pd.Size) if getattr(pd, "Size", None) else None
                except Exception:
                    size_val = None
                size_str = f"{size_val / (1024**3):.1f} GB" if size_val is not None else "N/A"
                # MediaType mapping: 3=HDD, 4=SSD, 5=SCM (treat as SSD)
                disk_type = "N/A"
                try:
                    mt = int(getattr(pd, "MediaType", 0))
                    if mt == 3:
                        disk_type = "HDD"
                    elif mt in (4, 5):
                        disk_type = "SSD"
                except Exception:
                    pass
                # Fallback hint via SpindleSpeed (0 -> SSD, >0 -> HDD)
                if disk_type == "N/A":
                    try:
                        sp = int(getattr(pd, "SpindleSpeed", 0))
                        if sp > 0:
                            disk_type = "HDD"
                        elif sp == 0:
                            disk_type = "SSD"
                    except Exception:
                        pass
                name = (getattr(pd, "FriendlyName", None) or getattr(pd, "Model", None) or "N/A").strip()
                disk_info.append({"Model": name, "Size": size_str, "Type": disk_type})
            if disk_info:
                return disk_info
        except Exception:
            # 2) Fallback: PowerShell Get-PhysicalDisk (Storage module)
            try:
                ps = 'powershell -NoProfile -Command "Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size | ConvertTo-Json -Compress"'
                out = subprocess.check_output(ps, shell=True, text=True, stderr=subprocess.DEVNULL)
                items = json.loads(out) if out and out.strip() else []
                if isinstance(items, dict):
                    items = [items]
                for it in items:
                    name = (it.get("FriendlyName") or "N/A").strip()
                    size_val = it.get("Size")
                    try:
                        size_val = int(size_val) if size_val is not None else None
                    except Exception:
                        size_val = None
                    size_str = f"{size_val / (1024**3):.1f} GB" if size_val is not None else "N/A"
                    mt = (it.get("MediaType") or "").strip().upper()
                    if mt in {"SSD", "SCM"}:
                        disk_type = "SSD"
                    elif mt == "HDD":
                        disk_type = "HDD"
                    else:
                        disk_type = "N/A"
                    disk_info.append({"Model": name, "Size": size_str, "Type": disk_type})
                if disk_info:
                    return disk_info
            except Exception:
                pass
        # 3) Last resort: Win32_DiskDrive with NVMe/SSD heuristics
        try:
            c = wmi.WMI()
            for disk in c.Win32_DiskDrive():
                if disk.Size:
                    size_gb = int(disk.Size) / (1024**3)
                    model = (disk.Model or "").strip()
                    pnp = (getattr(disk, "PNPDeviceID", "") or "").upper()
                    serial = (getattr(disk, "SerialNumber", "") or "").upper()
                    mdl_u = model.upper()
                    # Heuristics: NVMe implies SSD; model containing SSD also implies SSD
                    if "NVME" in pnp or "NVME" in mdl_u or "NVME" in serial:
                        disk_type = "SSD"
                    elif "SSD" in mdl_u:
                        disk_type = "SSD"
                    else:
                        disk_type = "HDD"
                    disk_info.append({"Model": model or "N/A", "Size": f"{size_gb:.1f} GB", "Type": disk_type})
        except Exception:
            pass
    else:
        try:
            result = subprocess.check_output("lsblk -d -o NAME,MODEL,SIZE,ROTA", shell=True, text=True)
            lines = result.split("\n")[1:]  # Skip header
            for line in lines:
                if line.strip():
                    parts = line.split()
                    if len(parts) >= 3:
                        model = " ".join(parts[1:-2])
                        size = parts[-2]
                        disk_type = "SSD" if parts[-1] == "0" else "HDD"
                        disk_info.append({"Model": model, "Size": size, "Type": disk_type})
        except Exception:
            pass
    return disk_info

def generate_html_report(system_info, cpu_info, ram_info, gpu_info, disk_info, config=None):
    # Load config if not provided
    if config is None:
        try:
            config = load_config()
        except Exception:
            config = {}

    title = config.get("title", "Gaming PC")

    # Build optional motherboard debug section
    debug_rows = [(k, v) for k, v in system_info.items() if str(k).startswith("Debug:")]
    debug_html = ""
    if debug_rows:
        debug_html = (
            '<div class="section">'
            '<h2>Debug: Motherboard Sources</h2>'
            '<table>'
            + ''.join(f'<tr><td>{k}</td><td>{v}</td></tr>' for k, v in debug_rows)
            + '</table>'
            '</div>'
        )
    # Build optional GPU debug section
    gpu_debug_html = ""
    try:
        rows = list(GPU_DEBUG)
    except Exception:
        rows = []
    if rows:
        gpu_debug_html = (
            '<div class="section">'
            '<h2>Debug: GPU Sources</h2>'
            '<table>'
            + ''.join(f'<tr><td style="color:#888;">{i+1:02d}</td><td>{line}</td></tr>' for i, line in enumerate(rows))
            + '</table>'
            '</div>'
        )
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>{title} Specifications</title>
        <style>
            body {{ 
                font-family: Arial, sans-serif; 
                margin: 40px; 
                background-color: #1a1a1a; 
                color: #ffffff; 
            }}
            .container {{ max-width: 1000px; margin: 0 auto; }}
            .section {{ 
                background-color: #2d2d2d; 
                padding: 20px; 
                margin-bottom: 20px; 
                border-radius: 8px; 
            }}
            h1 {{ color: #00ff9d; text-align: center; }}
            h2 {{ color: #00ccff; border-bottom: 2px solid #3d3d3d; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            td, th {{ padding: 12px; text-align: left; border-bottom: 1px solid #3d3d3d; }}
            th {{ background-color: #333333; }}
            .footer {{ text-align: center; margin-top: 30px; color: #888; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🎮 {title} Specifications 🖥️</h1>
            <div class="section">
                <h2>System Information</h2>
                <table>
                    {''.join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in system_info.items() if not str(key).startswith('Debug:'))}
                </table>
            </div>

            <div class="section">
                <h2>Processor (CPU)</h2>
                <table>
                    {''.join(f'<tr><td>{key}</td><td>{value}</td></tr>' for key, value in cpu_info.items())}
                </table>
            </div>

            <div class="section">
                <h2>Memory (RAM) - Total: {ram_info['Total RAM']}</h2>
                <table>
                    <tr><th>Module</th><th>Manufacturer</th><th>Capacity</th><th>Speed</th><th>Part Number</th></tr>
                    {''.join(
                        f'<tr><td>{m["Module"]}</td><td>{m["Manufacturer"]}</td><td>{m["Capacity"]}</td>'
                        f'<td>{m["Speed"]}</td><td>{m["Part Number"]}</td></tr>' 
                        for m in ram_info["Modules"])
                    }
                </table>
            </div>

            <div class="section">
                <h2>Graphics Card (GPU)</h2>
                <table>
                    <tr><th>Model</th><th>VRAM</th></tr>
                    {''.join(f'<tr><td>{g["Model"]}</td><td>{g["VRAM"]}</td></tr>' for g in gpu_info)}
                </table>
            </div>

            <div class="section">
                <h2>Storage Devices</h2>
                <table>
                    <tr><th>Model</th><th>Size</th><th>Type</th></tr>
                    {''.join(f'<tr><td>{d["Model"]}</td><td>{d["Size"]}</td><td>{d["Type"]}</td></tr>' for d in disk_info)}
                </table>
            </div>

            {debug_html}
            {gpu_debug_html}

            <div class="footer">
                Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            </div>
        </div>
    </body>
    </html>
    """
    return html

# ---------- Helpers for condensed share card ----------
def _short_vendor(name: str) -> str:
    s = (name or "").strip()
    u = s.upper()
    if "MICRO-STAR" in u or u.startswith("MSI") or u.startswith("MS-"):
        return "MSI"
    if "ASUSTEK" in u or "ASUS" in u:
        return "ASUS"
    if "GIGABYTE" in u:
        return "Gigabyte"
    if "ASROCK" in u:
        return "ASRock"
    if "LENOVO" in u:
        return "Lenovo"
    if "HEWLETT-PACKARD" in u or u.startswith("HP"):
        return "HP"
    if "DELL" in u:
        return "Dell"
    return s or "N/A"
def _clean_cpu_model(name: str) -> str:
    if not name:
        return "N/A"
    s = name
    s = s.replace("(R)", "").replace("(TM)", "").replace("CPU", "").replace("  ", " ")
    s = re.sub(r"with Radeon Graphics", "", s, flags=re.I)
    return s.strip()

def _parse_gb(text: str) -> float:
    if not text:
        return 0.0
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*GB", str(text), re.I)
    return float(m.group(1)) if m else 0.0

def _fmt_gb_from_bytes(b) -> str:
    try:
        n = int(b)
    except Exception:
        return "N/A"
    if n < 0:
        # Clamp negatives (some providers may return signed overflow)
        n = 0
    gb = n / (1024**3)
    # Avoid "-0.0" by clamping tiny values to 0
    if abs(gb) < 0.05:
        gb = 0.0
    return f"{gb:.1f} GB"

def _as_rgb(val, default):
    try:
        if isinstance(val, (list, tuple)) and len(val) >= 3:
            return (int(val[0]), int(val[1]), int(val[2]))
        if isinstance(val, str):
            m = re.match(r"#?([0-9a-fA-F]{6})", val)
            if m:
                hexv = m.group(1)
                return (int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16))
    except Exception:
        pass
    return default

def _deep_update(base: dict, override: dict) -> dict:
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            base[k] = _deep_update(base[k], v)
        else:
            base[k] = v
    return base

def _default_config() -> dict:
    return {
        "image_size": [1200, 675],
        "background_image": "bg.jpg",
        "background_fit": "cover",  # cover | contain | stretch
        "background_overlay": {"color": [0, 0, 0], "opacity": 0.4},
        "background_color": [26, 26, 26],
        "accent_color": [0, 255, 157],
        "sub_color": [0, 204, 255],
        "text_color": [240, 240, 240],
        "dim_color": [160, 160, 160],
        "font_paths": {"title": "", "h2": "", "body": "", "small": ""},
        "font_sizes": {"title": 56, "h2": 36, "body": 28, "small": 24},
    }

def load_config(config_filename: str = "xpec.config.json") -> dict:
    cfg = _default_config()

    # First: detect if running as a PyInstaller exe
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)  # folder where xpec.exe is
    else:
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base_dir = os.getcwd()

    # Candidate paths: next to exe/py, then current working dir
    paths = [
        os.path.join(base_dir, config_filename),
        os.path.join(os.getcwd(), config_filename)
    ]

    user_cfg = None
    for path in paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    user_cfg = json.load(f)
                break
            except Exception:
                pass

    if user_cfg:
        cfg = _deep_update(cfg, user_cfg)
    else:
        # Write a default config next to exe if none found
        try:
            with open(paths[0], "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    return cfg


def _load_font(path: str, size: int):
    try:
        if path and os.path.exists(path):
            return ImageFont.truetype(path, size)
    except Exception:
        pass
    try:
        # Windows Arial fallback if available
        win_font = r"C:\\Windows\\Fonts\\arial.ttf"
        if os.name == "nt" and os.path.exists(win_font):
            return ImageFont.truetype(win_font, size)
    except Exception:
        pass
    return ImageFont.load_default()

def _apply_background(img_path: str, size: tuple, bg_color: tuple, fit: str, overlay_color: tuple, overlay_opacity: float):
    W, H = size
    try:
        src = Image.open(img_path).convert("RGB")
    except Exception:
        return Image.new("RGB", (W, H), bg_color)
    sw, sh = src.size
    fit_mode = str(fit or "cover").lower()
    resample = getattr(Image, "Resampling", Image).LANCZOS
    if fit_mode == "stretch":
        canvas = src.resize((W, H), resample=resample)
    elif fit_mode == "contain":
        scale = min(W / sw, H / sh) if sw and sh else 1.0
        new_size = (max(1, int(sw * scale)), max(1, int(sh * scale)))
        resized = src.resize(new_size, resample=resample)
        canvas = Image.new("RGB", (W, H), bg_color)
        px = (W - new_size[0]) // 2
        py = (H - new_size[1]) // 2
        canvas.paste(resized, (px, py))
    else:  # cover
        scale = max(W / sw, H / sh) if sw and sh else 1.0
        new_size = (max(1, int(sw * scale)), max(1, int(sh * scale)))
        resized = src.resize(new_size, resample=resample)
        # center crop
        left = (new_size[0] - W) // 2
        top = (new_size[1] - H) // 2
        canvas = resized.crop((left, top, left + W, top + H))
    # Overlay
    try:
        if overlay_opacity and overlay_opacity > 0:
            overlay = Image.new("RGBA", (W, H), overlay_color + (int(255 * min(1.0, max(0.0, overlay_opacity))),))
            canvas = Image.alpha_composite(canvas.convert("RGBA"), overlay).convert("RGB")
    except Exception:
        pass
    return canvas

def _summarize_ram(ram_info: dict) -> str:
    total = ram_info.get("Total RAM", "N/A")
    modules = ram_info.get("Modules", [])
    count = len(modules)
    # derive per-module size and speed if consistent
    sizes = []
    speeds = []
    for m in modules:
        sizes.append(_parse_gb(m.get("Capacity", "")))
        sp = m.get("Speed", "")
        sm = re.search(r"(\d+)", sp)
        speeds.append(int(sm.group(1)) if sm else None)
    uniq_speed = None
    filtered = [s for s in speeds if s]
    if filtered and len(set(filtered)) == 1:
        uniq_speed = filtered[0]
    per_size = None
    if count and sizes and all(x > 0 for x in sizes):
        if len(set(sizes)) == 1:
            per_size = sizes[0]
    parts = [f"Total: {total}"]
    if count:
        if per_size:
            parts.append(f"({count}x{per_size:.0f} GB)")
        else:
            parts.append(f"({count} modules)")
    if uniq_speed:
        parts.append(f"@ {uniq_speed} MHz")
    return " ".join(parts)

def _choose_primary_gpu(gpu_info: list) -> dict:
    if not gpu_info:
        return {"Model": "N/A", "VRAM": "N/A"}
    return max(gpu_info, key=lambda g: _parse_gb(g.get("VRAM", "0 GB")))

def _summarize_storage(disk_info: list) -> str:
    if not disk_info:
        return "N/A"
    ssd = [d for d in disk_info if d.get("Type") == "SSD"]
    hdd = [d for d in disk_info if d.get("Type") == "HDD"]
    def total_gb(items):
        tot = 0.0
        for it in items:
            tot += _parse_gb(it.get("Size", "0 GB"))
        return tot
    parts = []
    if ssd:
        parts.append(f"{len(ssd)}x SSD ({total_gb(ssd):.1f} GB)")
    if hdd:
        parts.append(f"{len(hdd)}x HDD ({total_gb(hdd):.1f} GB)")
    return ", ".join(parts)

def generate_share_image(system_info, cpu_info, ram_info, gpu_info, disk_info, outfile="pc_specs.png", config=None):
    if not PIL_AVAILABLE:
        return None
    # Load config if not provided
    if config is None:
        try:
            config = load_config()
        except Exception:
            config = {}
    size = config.get("image_size", [1200, 675])
    try:
        W, H = int(size[0]), int(size[1])
    except Exception:
        W, H = 1200, 675
    bg_color = _as_rgb(config.get("background_color", [26, 26, 26]), (26, 26, 26))
    accent = _as_rgb(config.get("accent_color", [0, 255, 157]), (0, 255, 157))
    sub = _as_rgb(config.get("sub_color", [0, 204, 255]), (0, 204, 255))
    text = _as_rgb(config.get("text_color", [240, 240, 240]), (240, 240, 240))
    dim = _as_rgb(config.get("dim_color", [160, 160, 160]), (160, 160, 160))

    img = Image.new("RGB", (W, H), bg_color)
    bg_path = str(config.get("background_image", "") or "bg.jpg").strip()
    fit = config.get("background_fit", "cover")
    overlay_cfg = config.get("background_overlay", {})
    overlay_color = _as_rgb(overlay_cfg.get("color", [0, 0, 0]), (0, 0, 0))
    try:
        overlay_opacity = float(overlay_cfg.get("opacity", 0.4))
    except Exception:
        overlay_opacity = 0.0
    if bg_path and os.path.exists(bg_path):
        try:
            img = _apply_background(bg_path, (W, H), bg_color, fit, overlay_color, overlay_opacity)
        except Exception:
            pass
    elif overlay_opacity and overlay_opacity > 0:
        try:
            overlay = Image.new("RGBA", (W, H), overlay_color + (int(255 * min(1.0, max(0.0, overlay_opacity))),))
            img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
        except Exception:
            pass

    draw = ImageDraw.Draw(img)
    # Fonts
    title = config.get("title", "Gaming PC")
    font_paths = config.get("font_paths", {})
    font_sizes = config.get("font_sizes", {})
    tsize = int(font_sizes.get("title", 56) or 56)
    h2size = int(font_sizes.get("h2", 36) or 36)
    bodysize = int(font_sizes.get("body", 28) or 28)
    smallsize = int(font_sizes.get("small", 24) or 24)
    title_font = _load_font(font_paths.get("title", ""), tsize)
    h2_font = _load_font(font_paths.get("h2", ""), h2size)
    body_font = _load_font(font_paths.get("body", ""), bodysize)
    small_font = _load_font(font_paths.get("small", ""), smallsize)

    # Content
    mobo = system_info.get("Motherboard", "N/A")
    os_str = system_info.get("OS", "N/A")
    cpu_line = f"{_clean_cpu_model(cpu_info.get('Model', ''))}  |  {cpu_info.get('Cores', 'N/A')}C/{cpu_info.get('Threads', 'N/A')}T  |  {cpu_info.get('Max Clock', 'N/A')}"
    ram_line = _summarize_ram(ram_info)
    gpu_primary = _choose_primary_gpu(gpu_info)
    gpu_line = f"{gpu_primary.get('Model', 'N/A')}  |  {gpu_primary.get('VRAM', 'N/A')} VRAM"
    storage_line = _summarize_storage(disk_info)

    # Layout
    x = 60
    y = 50
    draw.text((x, y), title, fill=accent, font=title_font)
    y += 80
    draw.text((x, y), "GPU", fill=sub, font=h2_font); y += 48
    draw.text((x, y), gpu_line, fill=text, font=body_font); y += 50
    draw.text((x, y), "CPU", fill=sub, font=h2_font); y += 48
    draw.text((x, y), cpu_line, fill=text, font=body_font); y += 50
    draw.text((x, y), "RAM", fill=sub, font=h2_font); y += 48
    draw.text((x, y), ram_line, fill=text, font=body_font); y += 50
    draw.text((x, y), "Motherboard", fill=sub, font=h2_font); y += 48
    draw.text((x, y), mobo, fill=text, font=body_font); y += 50
    draw.text((x, y), "Storage", fill=sub, font=h2_font); y += 48
    draw.text((x, y), storage_line, fill=text, font=body_font); y += 50
    # Footer
    footer = f"{os_str}  •  Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    draw.text((x, H - 60), footer, fill=dim, font=small_font)

    img.save(outfile)
    return outfile

def main():
    # Load config and generate shareable PNG card
    try:
        cfg = load_config()
    except Exception:
        cfg = None

    try:
        system_info = get_system_info()
        cpu_info = get_cpu_info()
        ram_info = get_ram_info()
        gpu_info = get_gpu_info()
        disk_info = get_disk_info()

        html_content = generate_html_report(system_info, cpu_info, ram_info, gpu_info, disk_info, config=cfg)
        
        filename = "pc_specs.html"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        webbrowser.open(filename)
        
        try:
            png_path = generate_share_image(system_info, cpu_info, ram_info, gpu_info, disk_info, outfile="pc_specs.png", config=cfg)
            if png_path and os.name == "nt":
                try:
                    os.startfile(png_path)
                except Exception:
                    pass
        except Exception:
            pass
        
    except Exception as e:
        print(f"Error: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
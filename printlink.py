from flask import Flask, request, jsonify, render_template_string
import win32print
import win32api
import tempfile
import os
import base64
import requests
import hashlib
import struct
import threading
import time
import subprocess
import sys
from datetime import datetime
from PIL import Image, ImageOps
import io
import winreg
import signal
import atexit

app = Flask(__name__)

# ===========================================
# 🔐 FIXED PASSWORD CONFIGURATION
# ===========================================
FIXED_PASSWORD = "mypass"  # ⚠️ CHANGE THIS!

# ===========================================
# 🔹 Configuration Storage (Windows Registry)
# ===========================================
REG_PATH = r"Software\PrintServer"

def get_config():
    """Read configuration from Windows Registry"""
    config = {
        "site": "",
        "provider": "http",
        "host": "",
        "port": "",
        "email": "",
        "start_vortex": "true"  # Default to true for backward compatibility
    }
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ)
        for name in config.keys():
            try:
                value, _ = winreg.QueryValueEx(key, name)
                config[name] = value
            except:
                pass
        winreg.CloseKey(key)
    except:
        pass
    return config

def save_config(config):
    """Save configuration to Windows Registry"""
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        for name, value in config.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False

def is_configured():
    """Check if all required fields are configured"""
    config = get_config()
    required = ["site", "provider", "host", "port", "email"]
    return all(config.get(k) for k in required)

def should_start_vortex():
    """Check if vortex should be started based on configuration"""
    config = get_config()
    return config.get("start_vortex", "true").lower() == "true"

# ===========================================
# 🔹 Vortex Status Tracking
# ===========================================
vortex_status = {
    "running": False,
    "last_start": None,
    "last_error": None,
    "restart_count": 0,
    "process_id": None,
    "last_output": []
}

vortex_process = None
flask_shutdown = False

def update_vortex_status(running=None, error=None, pid=None, output=None):
    """Update vortex status"""
    global vortex_status
    if running is not None:
        vortex_status["running"] = running
        if running:
            vortex_status["last_start"] = datetime.now().isoformat()
            vortex_status["restart_count"] += 1
    if error is not None:
        vortex_status["last_error"] = error
    if pid is not None:
        vortex_status["process_id"] = pid
    if output is not None:
        vortex_status["last_output"].append(output)
        vortex_status["last_output"] = vortex_status["last_output"][-50:]

# ===========================================
# 🔹 Service Control
# ===========================================
def stop_all_services():
    """Stop vortex and prepare for shutdown"""
    global vortex_process, flask_shutdown
    flask_shutdown = True
    
    print("\n" + "=" * 50)
    print("🛑 Shutting down Print Server...")
    print("=" * 50)
    
    if vortex_process:
        try:
            print("   Stopping vortex process...")
            vortex_process.terminate()
            vortex_process.wait(timeout=5)
            print("   ✓ Vortex stopped")
        except:
            try:
                vortex_process.kill()
                print("   ✓ Vortex force stopped")
            except:
                pass
    
    update_vortex_status(running=False, error="Service stopped by user")
    print("   ✓ All services stopped")
    print("=" * 50 + "\n")

# Register cleanup handlers
atexit.register(stop_all_services)

def signal_handler(signum, frame):
    """Handle system signals"""
    stop_all_services()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# ===========================================
# 🔹 Find vortex.exe
# ===========================================
def find_vortex():
    """Find vortex.exe in the same directory as executable/script or bundled inside"""
    search_paths = []
    
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        search_paths.append(os.path.join(exe_dir, "vortex.exe"))
        
        if hasattr(sys, '_MEIPASS'):
            search_paths.append(os.path.join(sys._MEIPASS, "vortex.exe"))
        
        search_paths.append(os.path.join(os.getcwd(), "vortex.exe"))
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        search_paths.append(os.path.join(script_dir, "vortex.exe"))
        search_paths.append(os.path.join(os.getcwd(), "vortex.exe"))
    
    for vortex_path in search_paths:
        if os.path.exists(vortex_path):
            print(f"✓ Found vortex.exe at: {vortex_path}")
            return vortex_path
    
    print(f"❌ vortex.exe not found in any of these locations:")
    for path in search_paths:
        print(f"   - {path}")
    return None

# ===========================================
# 🔹 Professional Fluent Design UI
# ===========================================


# The Python variable CONFIG_HTML contains this new, optimized HTML:
CONFIG_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Print Server Configuration</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');

        :root {
            --primary: #0078D4; /* Microsoft Blue */
            --primary-dark: #005EA2;
            --success: #107C10; /* Dark Green */
            --error: #D83B01; /* Orange Red */
            --bg-light: #F3F5F9;
            --text-dark: #222;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: "Segoe UI", sans-serif;
            background-color: var(--bg-light);
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            color: var(--text-dark);
        }

        .container {
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.08), 0 0 0 1px rgba(0,0,0,0.04);
            width: 90%;
            max-width: 480px;
            padding: 30px;
        }

        .header {
            text-align: center;
            margin-bottom: 25px;
        }

        .logo {
            font-size: 40px;
            line-height: 1;
            margin-bottom: 5px;
        }

        h1 {
            font-size: 24px;
            font-weight: 600;
            color: var(--text-dark);
        }

        .subtitle {
            font-size: 14px;
            color: #777;
        }

        .alert {
            padding: 12px 15px;
            border-radius: 6px;
            font-size: 13px;
            margin-bottom: 20px;
            display: none;
            font-weight: 500;
        }

        .alert-success {
            background: #EAF7E9;
            color: var(--success);
            border-left: 4px solid var(--success);
        }

        .alert-error {
            background: #FEEEEE;
            color: var(--error);
            border-left: 4px solid var(--error);
        }

        .form-group {
            margin-bottom: 20px;
        }

        label {
            display: block;
            margin-bottom: 5px;
            font-size: 13px;
            color: #555;
            font-weight: 600;
        }

        .input-wrapper {
            position: relative;
            display: flex;
            align-items: center;
        }

        .input-icon {
            position: absolute;
            left: 12px;
            color: #aaa;
            font-size: 16px;
        }

        input:not([type="checkbox"]) {
            width: 100%;
            padding: 10px 12px 10px 35px; /* Added left padding for icon */
            border: 1px solid #ddd;
            border-radius: 6px;
            font-family: inherit;
            font-size: 14px;
            transition: border-color 0.2s, box-shadow 0.2s;
        }

        input:focus:not([type="checkbox"]) {
            border-color: var(--primary);
            box-shadow: 0 0 0 2px rgba(0,120,212,0.15);
            outline: none;
        }

        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-top: 5px;
        }
        
        .checkbox-group label {
            font-size: 14px;
            font-weight: 400;
            margin-bottom: 0;
            color: var(--text-dark);
        }

        .checkbox-help {
            color: #888;
            font-size: 11px;
            margin-left: 25px;
            margin-top: 3px;
        }

        .button-group {
            display: flex;
            gap: 10px;
            margin-top: 30px;
            justify-content: space-between;
        }

        button {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            width: 50%;
        }

        .btn-primary {
            background: var(--primary);
            color: white;
        }

        .btn-primary:hover {
            background: var(--primary-dark);
        }

        .btn-danger {
            background: #EFEFEF;
            color: var(--error);
            border: 1px solid #ddd;
        }

        .btn-danger:hover {
            background: #E0E0E0;
        }

        .footer-links {
            text-align: center;
            padding-top: 20px;
            margin-top: 20px;
            border-top: 1px solid #eee;
        }

        .footer-links a {
            color: var(--primary);
            text-decoration: none;
            margin: 0 10px;
            font-size: 13px;
            transition: color 0.2s;
        }

        .footer-links a:hover {
            color: var(--primary-dark);
        }

    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="logo">🖨️</div>
            <h1>Print Server</h1>
            <p class="subtitle">Configuration Utility</p>
        </div>

        <div id="alert-success" class="alert alert-success"></div>
        <div id="alert-error" class="alert alert-error"></div>

        {% if configured %}
        <div class="alert alert-success" style="display: block;">
            <span style="font-size: 16px; margin-right: 5px;">✅</span> Configuration is **Active**. Changes require a restart.
        </div>
        {% endif %}

        <form id="configForm">
            <div class="form-group">
                <label for="site">Site URL</label>
                <div class="input-wrapper">
                    <span class="input-icon">🌐</span>
                    <input type="text" name="site" id="site" value="{{ config.site }}" placeholder="https://your-site.com" required>
                </div>
            </div>

            <input type="hidden" name="provider" value="http" required>

            <div class="form-group">
                <label for="host">Host Address</label>
                <div class="input-wrapper">
                    <span class="input-icon">🖥️</span>
                    <input type="text" name="host" id="host" value="{{ config.host }}" placeholder="0.0.0.0" required>
                </div>
            </div>

            <div class="form-group">
                <label for="port">Port</label>
                <div class="input-wrapper">
                    <span class="input-icon">🔌</span>
                    <input type="text" name="port" id="port" value="{{ config.port }}" placeholder="9100" required>
                </div>
            </div>

            <div class="form-group">
                <label for="email">Email</label>
                <div class="input-wrapper">
                    <span class="input-icon">📧</span>
                    <input type="email" name="email" id="email" value="{{ config.email }}" placeholder="admin@example.com" required>
                </div>
            </div>

            <div class="form-group">
                <div class="checkbox-group">
                    <input type="checkbox" name="start_vortex" id="start_vortex" style="width: auto;" {{ 'checked' if config.get('start_vortex', 'true') == 'true' else '' }}>
                    <label for="start_vortex">Start Vortex Service</label>
                </div>
                <div class="checkbox-help">
                    Controls background service for remote printing. Disable if using localhost only.
                </div>
            </div>

            <div class="button-group">
                <button type="submit" class="btn-primary">💾 Save Settings</button>
                <button type="button" class="btn-danger" onclick="stopService()">🛑 Stop Service</button>
            </div>
        </form>

        <div class="footer-links">
            <a href="/status" target="_blank">📊 Status</a>
            <a href="/" target="_blank">📋 Printers</a>
            <a href="/api/docs" target="_blank">📚 API Docs</a>
        </div>
    </div>

    <script>
        // Existing JavaScript logic remains, ensuring functionality
        document.getElementById('configForm').onsubmit = async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = Object.fromEntries(formData);

            // Convert checkbox value to string
            data.start_vortex = document.getElementById('start_vortex').checked ? 'true' : 'false';

            const successAlert = document.getElementById('alert-success');
            const errorAlert = document.getElementById('alert-error');
            successAlert.style.display = 'none';
            errorAlert.style.display = 'none';

            try {
                const res = await fetch('/config', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(data)
                });
                const result = await res.json();

                if (result.success) {
                    successAlert.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">✓</span> Configuration saved successfully! **Restarting...**';
                    successAlert.style.display = 'block';
                    setTimeout(() => location.reload(), 2000);
                } else {
                    throw new Error(result.error || 'Failed to save configuration');
                }
            } catch (err) {
                errorAlert.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">❌</span> ' + err.message;
                errorAlert.style.display = 'block';
            }
        };

        async function stopService() {
            if (!confirm('🛑 Are you sure you want to stop ALL services? This will require manual restart of the application.')) return;

            const successAlert = document.getElementById('alert-success');
            const errorAlert = document.getElementById('alert-error');
            successAlert.style.display = 'none';
            errorAlert.style.display = 'none';

            try {
                const res = await fetch('/shutdown', { method: 'POST' });
                const result = await res.json();

                if (result.success) {
                    successAlert.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">✓</span> Services stopped successfully. **Application is closing shortly.**';
                    successAlert.style.display = 'block';
                } else {
                    throw new Error(result.error || 'Failed to stop services');
                }
            } catch (err) {
                errorAlert.innerHTML = '<span style="font-size: 16px; margin-right: 5px;">❌</span> ' + err.message;
                errorAlert.style.display = 'block';
            }
        }
    </script>
</body>
</html>
"""

# The Python variable STATUS_HTML contains this new, optimized HTML:
STATUS_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Vortex Status Monitor</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;500&family=Segoe+UI:wght@400;600;700&display=swap');

        :root {
            --primary: #0078D4;
            --running: #4CAF50;
            --stopped: #F44336;
            --log-bg: #282C34; /* Dark background for logs */
            --log-text: #ABB2BF;
            --card-bg: #FFFFFF;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', sans-serif;
            color: #222;
            padding: 20px;
            background-color: #F8F8FC;
            min-height: 100vh;
        }

        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: var(--card-bg);
            border-radius: 12px;
            padding: 30px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }

        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
        }

        h1 {
            font-size: 26px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .controls button {
            padding: 8px 18px;
            border: none;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }

        .btn-refresh {
            background: var(--primary);
            color: white;
        }

        .btn-refresh:hover {
            background: #005EA2;
        }
        
        .btn-back {
            background: #EBEBEB;
            color: #444;
        }
        
        .btn-back:hover {
            background: #DDD;
        }

        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            padding: 10px 20px;
            border-radius: 20px;
            font-size: 15px;
            font-weight: 600;
            color: white;
            margin-bottom: 25px;
        }

        .status-running {
            background: var(--running);
            box-shadow: 0 4px 8px rgba(76, 175, 80, 0.3);
        }

        .status-stopped {
            background: var(--stopped);
            box-shadow: 0 4px 8px rgba(244, 67, 54, 0.3);
        }

        .pulse-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: white;
        }
        
        .status-running .pulse-dot {
            animation: pulseDot 2s infinite;
        }

        @keyframes pulseDot {
            0%, 100% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.7; transform: scale(1.3); }
        }

        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .info-card {
            background: #F8F9FA;
            border-radius: 8px;
            padding: 15px 18px;
            border: 1px solid #E0E0E0;
        }

        .info-label {
            color: #777;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            margin-bottom: 5px;
        }

        .info-value {
            color: var(--log-text);
            font-size: 15px;
            font-family: 'Roboto Mono', monospace;
            background-color: #EEE; /* Light background for value */
            padding: 3px 6px;
            border-radius: 4px;
            display: inline-block;
            color: #333;
        }

        .log-section h2 {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 10px;
        }

        .log-container {
            background: var(--log-bg);
            border-radius: 8px;
            padding: 15px;
            max-height: 400px;
            overflow-y: auto;
            font-family: 'Roboto Mono', monospace;
            font-size: 12px;
            line-height: 1.5;
            color: var(--log-text);
            white-space: pre-wrap;
        }

        .log-line {
            padding: 2px 0;
        }

        .log-error {
            color: #FF6347; /* Error color for logs */
            font-weight: 500;
        }
        
        .info-value.log-error {
            color: var(--stopped);
            background-color: #FEEEEE;
        }

        .log-empty {
            color: #777;
            font-style: italic;
            text-align: center;
            padding: 20px 0;
        }

        .footer {
            margin-top: 25px;
            padding-top: 15px;
            border-top: 1px solid #E0E0E0;
            text-align: center;
            color: #999;
            font-size: 12px;
        }

        @media (max-width: 600px) {
            .header {
                flex-direction: column;
                align-items: flex-start;
            }
            .controls {
                margin-top: 15px;
                width: 100%;
                justify-content: space-between;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Vortex Status Monitor</h1>
            <div class="controls">
                <button class="btn-refresh" onclick="refreshStatus()">↻ Refresh</button>
                <button class="btn-back" onclick="window.location.href='/config'">⚙️ Configuration</button>
            </div>
        </div>

        <div class="status-badge {{ 'status-running' if status.running else 'status-stopped' }}">
            <div class="pulse-dot"></div>
            {{ 'VORTEX SERVICE RUNNING' if status.running else 'VORTEX SERVICE STOPPED' }}
        </div>

        <div class="info-grid">
            <div class="info-card">
                <div class="info-label">Last Start Time</div>
                <div class="info-value">{{ status.last_start or 'N/A' }}</div>
            </div>

            <div class="info-card">
                <div class="info-label">Process ID</div>
                <div class="info-value">{{ status.process_id or 'N/A' }}</div>
            </div>

            <div class="info-card">
                <div class="info-label">Restart Count</div>
                <div class="info-value">{{ status.restart_count }}</div>
            </div>

            <div class="info-card">
                <div class="info-label">Last Error</div>
                <div class="info-value log-error">{{ status.last_error or 'None' }}</div>
            </div>
        </div>

        <div class="log-section">
            <h2>📜 Console Log (Last 50 Entries)</h2>
            <div class="log-container">
                {% if status.last_output %}
                    {% for line in status.last_output %}
                    <div class="log-line">{{ line }}</div>
                    {% endfor %}
                {% else %}
                    <div class="log-empty">No output captured yet...</div>
                {% endif %}
            </div>
        </div>

        <div class="footer">
            Auto-refreshing every 5 seconds.
        </div>
    </div>

    <script>
        function refreshStatus() { location.reload(); }
        setInterval(refreshStatus, 5000);
        // Scroll to bottom of log
        const logContainer = document.querySelector('.log-container');
        if (logContainer) logContainer.scrollTop = logContainer.scrollHeight;
    </script>
</body>
</html>
"""

@app.route("/status", methods=["GET"])
def status_page():
    """Show vortex status page"""
    return render_template_string(STATUS_HTML, status=vortex_status)

@app.route("/api/status", methods=["GET"])
def api_status():
    """API endpoint for vortex status"""
    return jsonify(vortex_status)

@app.route("/", methods=["GET"])
def config_page():
    """Show configuration page"""
    config = get_config()
    return render_template_string(CONFIG_HTML, config=config, configured=is_configured())

@app.route("/config", methods=["POST"])
def save_config_endpoint():
    """Save configuration endpoint"""
    try:
        data = request.get_json(force=True)
        if save_config(data):
            global vortex_restart_flag
            vortex_restart_flag = True
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "error": "Failed to save configuration"}), 500
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route("/shutdown", methods=["POST"])
def shutdown_service():
    """Shutdown endpoint to stop all services"""
    try:
        # Stop vortex in background thread
        threading.Thread(target=lambda: (time.sleep(1), stop_all_services(), os._exit(0)), daemon=True).start()
        return jsonify({"success": True, "message": "Services stopping..."})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# ===========================================
# 🔹 Utility: stable printer ID
# ===========================================
def make_printer_id(p):
    name = p.get("pPrinterName", "") or p.get("Name", "")
    port = p.get("pPortName", "") or ""
    driver = p.get("pDriverName", "") or ""
    unique_str = f"{name}|{port}|{driver}"
    return hashlib.md5(unique_str.encode("utf-8")).hexdigest()[:8]

# ===========================================
# 🔹 Temp cleanup helper
# ===========================================
def schedule_remove(path, delay_seconds=60):
    def _remove():
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    t = threading.Timer(delay_seconds, _remove)
    t.daemon = True
    t.start()

# ===========================================
# 🔹 Manual CORS headers
# ===========================================
@app.after_request
def after_request(response):
    response.headers.add("Access-Control-Allow-Origin", "*")
    response.headers.add("Access-Control-Allow-Headers", "Content-Type")
    response.headers.add("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    return response

# ===========================================
# 🔹 Convert Image to ESC/POS bytes
# ===========================================
def image_to_escpos_bytes(img_data, is_url=False):
    if is_url:
        if img_data.startswith("data:image"):
            header, b64data = img_data.split(",", 1)
            im = Image.open(io.BytesIO(base64.b64decode(b64data)))
        else:
            r = requests.get(img_data)
            r.raise_for_status()
            im = Image.open(io.BytesIO(r.content))
    else:
        im = Image.open(io.BytesIO(base64.b64decode(img_data)))

    if im.mode != "1":
        im = im.convert("1")

    if im.size[0] % 8:
        new_width = im.size[0] + (8 - im.size[0] % 8)
        im2 = Image.new("1", (new_width, im.size[1]), "white")
        im2.paste(im, (0, 0))
        im = im2

    im = ImageOps.invert(im.convert("L")).convert("1")

    width_bytes = int(im.size[0] / 8)
    height = im.size[1]
    header = b"\x1d\x76\x30\x00" + struct.pack("2B", width_bytes % 256, width_bytes // 256)
    header += struct.pack("2B", height % 256, height // 256)
    return header + im.tobytes()

# ===========================================
# 🔹 Combine Logo + Text (ESC/POS)
# ===========================================
def build_escpos_with_logo(logo_data, text, is_url=False):
    if logo_data.startswith("data:image"):
        logo_bytes = image_to_escpos_bytes(logo_data, is_url=True)
    else:
        logo_bytes = image_to_escpos_bytes(logo_data, is_url=is_url)

    commands = b"".join([
        b"\x1B\x40",
        b"\x1B\x61\x01",
        logo_bytes,
        b"\x1B\x61\x00",
        b"\n",
        text.encode("utf-8"),
        b"\n\n\n",
        b"\x1D\x56\x00"
    ])
    return base64.b64encode(commands).decode("utf-8")

# ===========================================
# 🔹 Printer list
# ===========================================
@app.route("/printers", methods=["GET"])
def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags, None, 2)
    default = win32print.GetDefaultPrinter()
    result = []
    for p in printers:
        info = {
            "Id": make_printer_id(p),
            "Name": p.get("pPrinterName", ""),
            "PortName": p.get("pPortName", ""),
            "DriverName": p.get("pDriverName", ""),
            "Location": p.get("pLocation", ""),
            "Comment": p.get("pComment", ""),
            "ShareName": p.get("pShareName", ""),
            "Status": p.get("Status", 0),
            "Attributes": p.get("Attributes", 0),
            "IsDefault": (p.get("pPrinterName", "") == default),
        }
        result.append(info)
    return jsonify(result)

# ===========================================
# 🔹 Resolve printer by ID
# ===========================================
def resolve_printer(identifier):
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags, None, 2)
    for p in printers:
        name = p.get("pPrinterName") or ""
        if identifier == name:
            return name
    for p in printers:
        if identifier == make_printer_id(p):
            return p.get("pPrinterName")
    raise ValueError("Printer not found.")

# ===========================================
# 🔹 Print helpers
# ===========================================
def _print_text(printer_name, text):
    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, ("TextJob", None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, text.encode("utf-8"))
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)

def _print_raw(printer_name, b64data):
    data = base64.b64decode(b64data)
    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, ("RawPrintJob", None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, data)
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)

def _print_file(printer_name, path):
    win32api.ShellExecute(0, "printto", path, f'"{printer_name}"', ".", 0)

def _print_pdf(printer_name, pdf_data):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    if pdf_data.startswith("http"):
        r = requests.get(pdf_data)
        tmp.write(r.content)
    else:
        tmp.write(base64.b64decode(pdf_data))
    tmp.close()
    _print_file(printer_name, tmp.name)
    schedule_remove(tmp.name)

def _print_image(printer_name, img_data):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
    if img_data.startswith("http"):
        r = requests.get(img_data)
        tmp.write(r.content)
    else:
        tmp.write(base64.b64decode(img_data))
    tmp.close()
    _print_file(printer_name, tmp.name)
    schedule_remove(tmp.name)

# ===========================================
# 🔹 Print endpoint
# ===========================================
@app.route("/print", methods=["POST"])
def print_job():
    try:
        data = request.get_json(force=True)
    except Exception:
        return jsonify({"error": "Invalid JSON"}), 400

    printer_id = data.get("printer")
    mode = data.get("mode", "text")
    content = data.get("data")
    logo = data.get("logo")
    logo_url = data.get("logo_url")

    if not printer_id or not content:
        return jsonify({"error": "Missing printer or data"}), 400

    try:
        printer_name = resolve_printer(printer_id)
    except Exception as e:
        return jsonify({"error": str(e)}), 404

    try:
        if mode == "text":
            _print_text(printer_name, content)
        elif mode == "raw":
            _print_raw(printer_name, content)
        elif mode == "pdf":
            _print_pdf(printer_name, content)
        elif mode == "image":
            _print_image(printer_name, content)
        elif mode == "logo_text":
            if not logo and not logo_url:
                return jsonify({"error": "Missing 'logo' or 'logo_url'"}), 400
            combined = build_escpos_with_logo(logo or logo_url, content, is_url=bool(logo_url))
            _print_raw(printer_name, combined)
        else:
            return jsonify({"error": "Invalid mode"}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify({"status": "ok", "printer": printer_name, "mode": mode})

# ===========================================
# 🔹 API Documentation Endpoint
# ===========================================


# Add this new variable to your Flask application code
DOCS_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Print Server API Documentation</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;500&family=Segoe+UI:wght@400;600;700&display=swap');

        :root {
            --primary: #0078D4;
            --bg-light: #F8F8FC;
            --text-dark: #222;
            --code-bg: #282C34;
            --code-text: #ABB2BF;
            --method-get: #4CAF50; /* Green */
            --method-post: #0078D4; /* Blue */
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', sans-serif;
            background-color: var(--bg-light);
            color: var(--text-dark);
            padding: 40px 20px;
            line-height: 1.6;
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            padding: 30px;
        }

        .header {
            margin-bottom: 30px;
            border-bottom: 2px solid #E0E0E0;
            padding-bottom: 15px;
        }

        h1 {
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 5px;
            color: var(--primary);
        }

        .endpoint-card {
            border: 1px solid #E0E0E0;
            border-radius: 8px;
            margin-bottom: 25px;
            overflow: hidden;
        }

        .endpoint-header {
            padding: 15px;
            background-color: #F7F7F7;
            display: flex;
            align-items: center;
            gap: 15px;
            border-bottom: 1px solid #E0E0E0;
        }

        .method-badge {
            font-size: 12px;
            font-weight: 700;
            padding: 5px 10px;
            border-radius: 4px;
            color: white;
            text-transform: uppercase;
        }

        .method-get { background-color: var(--method-get); }
        .method-post { background-color: var(--method-post); }

        .path {
            font-family: 'Roboto Mono', monospace;
            font-size: 16px;
            font-weight: 500;
            color: var(--text-dark);
        }

        .description {
            padding: 15px;
            font-size: 14px;
            color: #555;
            border-bottom: 1px dashed #E0E0E0;
        }

        .content-grid {
            display: grid;
            grid-template-columns: 1fr;
        }

        .content-section {
            padding: 15px;
            position: relative;
        }
        
        .content-section:first-child {
            border-right: 1px solid #E0E0E0;
        }
        
        h3 {
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 10px;
            color: var(--primary);
        }

        .code-block {
            background-color: var(--code-bg);
            color: var(--code-text);
            font-family: 'Roboto Mono', monospace;
            font-size: 13px;
            padding: 15px;
            border-radius: 6px;
            overflow-x: auto;
            position: relative;
        }

        .copy-btn {
            position: absolute;
            top: 5px;
            right: 5px;
            background: rgba(255, 255, 255, 0.1);
            color: var(--code-text);
            border: none;
            padding: 5px 10px;
            border-radius: 4px;
            font-size: 11px;
            cursor: pointer;
            transition: background 0.2s;
        }

        .copy-btn:hover {
            background: rgba(255, 255, 255, 0.2);
        }

        @media (min-width: 650px) {
            .content-grid {
                grid-template-columns: 1fr 1fr;
            }
            .content-section:first-child {
                border-right: 1px solid #E0E0E0;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Print Server API Documentation</h1>
            <p style="font-size: 14px; color: #777;">Endpoints for interacting with the local print service.</p>
        </div>

        <div class="endpoint-card">
            <div class="endpoint-header">
                <span class="method-badge method-get">GET</span>
                <span class="path">/printers</span>
            </div>
            <div class="description">
                Retrieve a list of all local and connected printers.
            </div>
            <div class="content-grid">
                <div class="content-section">
                    <h3>Example Response</h3>
                    <div class="code-block" id="code1">
                        [
  {
    "Id": "f4e5a9c0",
    "Name": "Printer Name (Share)",
    "PortName": "USB001",
    "DriverName": "HP Universal Printing PCL 6",
    "IsDefault": true,
    "Status": 0
  },
  // ... more printers
]
                        <button class="copy-btn" onclick="copyCode('code1', this)">Copy</button>
                    </div>
                </div>
                <div class="content-section">
                    <h3>Printer ID</h3>
                    <p style="font-family: 'Roboto Mono', monospace; font-size: 13px; color: var(--method-post);">{{ '{{Id}}' }}</p>
                    <p style="font-size: 13px; color: #555; margin-top: 5px;">
                        The **8-character MD5 hash** is the stable identifier you should use for the `/print` endpoint, as printer names can change.
                    </p>
                </div>
            </div>
        </div>

        <div class="endpoint-card">
            <div class="endpoint-header">
                <span class="method-badge method-post">POST</span>
                <span class="path">/print</span>
            </div>
            <div class="description">
                Send a print job to the specified printer. Use the 8-character **Id** from the list above.
            </div>
            <div class="content-grid">
                <div class="content-section">
                    <h3>Request Body (JSON)</h3>
                    <div class="code-block" id="code2">
                        {
  "type": "raw",
  "data": "b3MxMjM...", // Base64 encoded RAW printer commands (e.g., ESC/POS)
  "pages": "1-3,5", // Optional: Page range for file/PDF printing
  "copies": 1 // Optional
}
                        <button class="copy-btn" onclick="copyCode('code2', this)">Copy</button>
                    </div>
                </div>
                <div class="content-section">
                    <h3>Supported Types</h3>
                    <ul style="padding-left: 20px; font-size: 13px; color: #333;">
                        <li style="margin-bottom: 5px;"><code style="background-color: #EEE; padding: 2px 4px; border-radius: 3px;">raw</code>: Base64 encoded byte data (for thermal/label printers).</li>
                        <li style="margin-bottom: 5px;"><code style="background-color: #EEE; padding: 2px 4px; border-radius: 3px;">file</code>: Base64 encoded PDF or image data, or a URL to a file.</li>
                        <li><code style="background-color: #EEE; padding: 2px 4px; border-radius: 3px;">text</code>: Plain text (will be printed as a simple text document).</li>
                        <li><code style="background-color: #EEE; padding: 2px 4px; border-radius: 3px;">text</code>: Logo text (will be printed as a simple text document with logo send logo_url body param with base64 or image link).</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div class="endpoint-card">
            <div class="endpoint-header">
                <span class="method-badge method-get">GET</span>
                <span class="path">/api/status</span>
            </div>
            <div class="description">
                Get the current running status, PID, and last error of the internal Vortex service.
            </div>
            <div class="content-grid">
                <div class="content-section" style="border-right: none;">
                    <h3>Example Response</h3>
                    <div class="code-block" id="code3">
                        {
  "running": true,
  "last_start": "2025-11-02T18:00:00.000000",
  "last_error": null,
  "restart_count": 5,
  "process_id": 1234
}
                        <button class="copy-btn" onclick="copyCode('code3', this)">Copy</button>
                    </div>
                </div>
            </div>
        </div>


    </div>

    <script>
        function copyCode(elementId, button) {
            const codeElement = document.getElementById(elementId);
            const textToCopy = codeElement.innerText.trim();
            
            // Use the modern clipboard API
            navigator.clipboard.writeText(textToCopy).then(() => {
                button.textContent = 'Copied!';
                setTimeout(() => {
                    button.textContent = 'Copy';
                }, 2000);
            }).catch(err => {
                console.error('Could not copy text: ', err);
                button.textContent = 'Error!';
            });
        }
    </script>
</body>
</html>
"""
@app.route("/api/docs", methods=["GET"])
def api_docs():
    """API Documentation"""
    config = get_config()
    base_url = f"https://{config.get('site')}" if config.get('site') else "http://localhost:9100"
    

    return render_template_string(DOCS_HTML)

# ===========================================
# 🔹 Auto-run vortex (background thread)
# ===========================================
vortex_restart_flag = False

def run_vortex():
    global vortex_restart_flag, vortex_process, flask_shutdown
    
    # Check if vortex should be started
    if not should_start_vortex():
        print("⚠️ Vortex startup disabled in configuration")
        update_vortex_status(running=False, error="Vortex startup disabled")
        return
    
    vortex_path = find_vortex()
    if not vortex_path:
        update_vortex_status(running=False, error="vortex.exe not found")
        print("❌ Vortex.exe not found. Print server will run without vortex.")
        print("   Visit http://localhost:9100/config to configure when vortex.exe is available.")
        return

    while not flask_shutdown:
        # Check if vortex should be running
        if not should_start_vortex():
            update_vortex_status(running=False, error="Vortex disabled by configuration")
            time.sleep(5)
            continue
            
        # Wait for configuration
        while not is_configured() and not flask_shutdown:
            update_vortex_status(running=False, error="Waiting for configuration")
            print("⏳ Waiting for configuration... Visit http://localhost:9100/config")
            time.sleep(5)
        
        if flask_shutdown:
            break
            
        config = get_config()
        vortex_restart_flag = False
        
        cmd = [
            vortex_path,
            "--site", config.get("site", ""),
            "--provider", "http",
            "--host", config.get("host", ""),
            "--port", config.get("port", ""),
            "--email", config.get("email", ""),
            "--password", FIXED_PASSWORD
        ]

        try:
            print(f"[{datetime.now()}] Launching vortex...")
            print(f"Command: {' '.join(cmd[:10])}... [password hidden]\n")
            
            # Run vortex silently without command prompt
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            
            vortex_process = subprocess.Popen(cmd, shell=False, startupinfo=startupinfo)
            
            update_vortex_status(running=True, error=None, pid=vortex_process.pid)
            print(f"✓ Vortex started with PID: {vortex_process.pid} (running silently)\n")

            # Monitoring loop
            while not flask_shutdown:
                if vortex_restart_flag:
                    print("🔄 Configuration changed. Restarting vortex...")
                    update_vortex_status(running=False, error="Restarting due to config change")
                    try:
                        vortex_process.terminate()
                        vortex_process.wait(timeout=5)
                    except:
                        try:
                            vortex_process.kill()
                        except:
                            pass
                    break
                
                # Check if vortex should still be running
                if not should_start_vortex():
                    print("🔄 Vortex disabled in configuration. Stopping...")
                    update_vortex_status(running=False, error="Stopped by configuration")
                    try:
                        vortex_process.terminate()
                        vortex_process.wait(timeout=5)
                    except:
                        try:
                            vortex_process.kill()
                        except:
                            pass
                    break
                
                exit_code = vortex_process.poll()
                if exit_code is not None:
                    error_msg = f"vortex exited with code {exit_code}"
                    print(f"[{datetime.now()}] {error_msg}")
                    update_vortex_status(running=False, error=error_msg)
                    break
                
                time.sleep(1)
                
        except Exception as e:
            error_msg = f"Error running vortex: {e}"
            print(f"[{datetime.now()}] {error_msg}")
            update_vortex_status(running=False, error=error_msg)
        
        if not vortex_restart_flag and not flask_shutdown and should_start_vortex():
            print("Restarting vortex in 5 seconds...\n")
            time.sleep(5)

# ===========================================
# 🔹 Start both services
# ===========================================
if __name__ == "__main__":
    print("=" * 60)
    print("🖨️  Professional Print Server - Enterprise Edition")
    print("=" * 60)
    
    print(f"\n🔍 System Information:")
    print(f"   - Running as exe: {getattr(sys, 'frozen', False)}")
    if getattr(sys, 'frozen', False):
        print(f"   - Executable: {sys.executable}")
        print(f"   - Directory: {os.path.dirname(sys.executable)}")
        if hasattr(sys, '_MEIPASS'):
            print(f"   - Temp bundle: {sys._MEIPASS}")
    else:
        print(f"   - Script: {os.path.abspath(__file__)}")
    print(f"   - Working dir: {os.getcwd()}")
    print(f"   - Password: {'*' * len(FIXED_PASSWORD)} (configured in script)")
    
    print(f"\n🌐 Access Points:")
    print(f"   ⚙️  Configuration: http://localhost:9100/config")
    print(f"   📊 Status Monitor: http://localhost:9100/status")
    print(f"   📋 Printer List:   http://localhost:9100/")
    print(f"   📚 API Docs:       http://localhost:9100/api/docs")
    
    print(f"\n⚠️  Important Notes:")
    print(f"   - Flask server: 0.0.0.0:9100 (accepts external connections)")
    print(f"   - Set Host to '0.0.0.0' in config for external vortex access")
    print(f"   - Password is hardcoded in script for security")
    print(f"   - Use Stop Service button or Ctrl+C to shutdown")
    print("=" * 60 + "\n")
    
    # Start vortex monitoring thread
    threading.Thread(target=run_vortex, daemon=True).start()
    
    # Run Flask server
    try:
        app.run(host="0.0.0.0", port=9100, debug=False)
    except KeyboardInterrupt:
        print("\n\nReceived shutdown signal...")
        stop_all_services()
    except Exception as e:
        print(f"\nFlask error: {e}")
        stop_all_services()
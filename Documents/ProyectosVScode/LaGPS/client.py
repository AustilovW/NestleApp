import subprocess
import json
import sys
import time
import os

def run_client():
    # Start the server process
    # We use shell=True to ensure uv is found, but direct execution is better if possible.
    # Assuming uv is in path.
    process = subprocess.Popen(
        ["uv", "run", "main.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, # Capture stderr to avoid polluting stdout
        text=True,
        cwd=r"c:\Users\austi\Documents\ProyectosVScode\LaGPS"
    )

    # Helper to send message
    def send(msg):
        json_str = json.dumps(msg)
        process.stdin.write(json_str + "\n")
        process.stdin.flush()

    # 1. Initialize
    send({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "client", "version": "1.0"}
        }
    })

    # 2. Notify initialized
    send({
        "jsonrpc": "2.0",
        "method": "notifications/initialized"
    })

    # 3. Call tool to add expense
    send({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "agregar_gasto",
            "arguments": {
                "fecha": "2025-11-21",
                "categoria": "Almuerzo",
                "cantidad": 25.0,
                "metodo_de_pago": "Tarjeta de Crédito"
            }
        }
    })

    # Read output loop
    start_time = time.time()
    while time.time() - start_time < 10:
        line = process.stdout.readline()
        if not line:
            break
        try:
            data = json.loads(line)
            if data.get("id") == 2 and "result" in data:
                print("Server response:", json.dumps(data["result"], indent=2))
                break
        except json.JSONDecodeError:
            continue
            
    process.terminate()

if __name__ == "__main__":
    run_client()

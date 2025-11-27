from fastmcp import FastMCP
import csv
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

# Constantes
GASTOS_FILE = Path(__file__).parent / 'gastos.csv'
MCP_NAME = 'Gastos MCP'

mcp = FastMCP(MCP_NAME)

def _inicializar_csv():
    """Crea el archivo CSV con encabezados si no existe."""
    if not GASTOS_FILE.exists():
        with GASTOS_FILE.open(mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['fecha', 'categoria', 'cantidad', 'metodo_de_pago'])

def _append_gasto(fila: List[Any]) -> None:
    """Escribe una nueva fila en el CSV."""
    _inicializar_csv() # Asegura que exista antes de escribir
    with GASTOS_FILE.open(mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(fila)

@mcp.tool
def agregar_gasto(fecha: str, categoria: str, cantidad: float, metodo_de_pago: str) -> str:
    """
    Registra un nuevo gasto en el sistema.
    
    Args:
        fecha: La fecha del gasto en formato YYYY-MM-DD.
        categoria: La categoría del gasto (ej. Comida, Transporte).
        cantidad: El monto gastado (debe ser positivo).
        metodo_de_pago: Cómo se pagó (ej. Efectivo, Tarjeta).
    """
    errores = []
    
    # Validar fecha
    try:
        datetime.strptime(fecha, "%Y-%m-%d")
    except ValueError:
        errores.append("Fecha inválida. Usa formato YYYY-MM-DD.")

    # Validar cantidad
    if cantidad <= 0:
        errores.append("La cantidad debe ser mayor a 0.")

    # Validar textos
    if not categoria.strip():
        errores.append("La categoría no puede estar vacía.")
    if not metodo_de_pago.strip():
        errores.append("El método de pago no puede estar vacío.")

    if errores:
        return f"Error al agregar gasto: {'; '.join(errores)}"

    try:
        _append_gasto([fecha, categoria, cantidad, metodo_de_pago])
        return f"Gasto agregado exitosamente: {categoria} - ${cantidad}"
    except Exception as e:
        return f"Error crítico al guardar el gasto: {e}"

@mcp.resource('resource://gastos')
def datos_de_gastos() -> Dict[str, Any]:
    """Lee y devuelve todos los gastos registrados."""
    if not GASTOS_FILE.exists():
        return {"gastos": [], "mensaje": "No hay gastos registrados."}
    
    try:
        with GASTOS_FILE.open(mode='r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            columnas = reader.fieldnames if reader.fieldnames else []
            gastos = []
            
            for row in reader:
                gasto = dict(row)
                # Formateo de fecha para visualización
                if 'fecha' in gasto:
                    try:
                        dt = datetime.strptime(gasto['fecha'], "%Y-%m-%d")
                        gasto['fecha_formateada'] = dt.strftime("%d/%m/%Y")
                    except ValueError:
                        gasto['fecha_formateada'] = gasto['fecha'] # Fallback
                
                # Convertir cantidad a float para que sea numérico en el JSON
                if 'cantidad' in gasto:
                    try:
                        gasto['cantidad'] = float(gasto['cantidad'])
                    except ValueError:
                        pass

                gastos.append(gasto)

        return {
            "columnas": columnas,
            "gastos": gastos,
            "total_registros": len(gastos)
        }
    except Exception as e:
        return {"error": f"No se pudieron leer los gastos: {e}"}

@mcp.prompt
def prompt_agregar_gasto() -> str:
    return 'Usa la herramienta agregar_gasto para registrar un nuevo gasto. Pide los detalles si faltan.'

if __name__ == "__main__":
    mcp.run(show_banner=False)

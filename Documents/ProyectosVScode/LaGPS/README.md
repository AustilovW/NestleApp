# LaGPS - Gestor de Gastos MCP

Este es un servidor MCP (Model Context Protocol) simple construido con Python y `fastmcp` para gestionar gastos personales en un archivo CSV.

## Requisitos

- Python 3.12+
- [uv](https://github.com/astral-sh/uv) (recomendado para gestión de paquetes)

## Instalación

1. Clona este repositorio.
2. Instala las dependencias:

```bash
uv sync
```

O si usas pip tradicional:

```bash
pip install fastmcp
```

## Uso

### Ejecutar el Servidor

Puedes ejecutar el servidor directamente con Python:

```bash
# Si usas uv
uv run main.py

# O con python directo (asegúrate de activar tu venv)
python main.py
```

### Probar con MCP Inspector

La mejor forma de probar tu servidor mientras desarrollas es usando el Inspector de MCP.

```bash
npx @modelcontextprotocol/inspector uv run main.py
```

Esto abrirá una interfaz web donde puedes ver tus herramientas (`agregar_gasto`) y recursos (`gastos`), y probarlos interactivamente.

## Herramientas Disponibles

### `agregar_gasto`
Registra un nuevo gasto en el sistema.
- **fecha**: YYYY-MM-DD
- **categoria**: Texto (ej. Comida)
- **cantidad**: Número
- **metodo_de_pago**: Texto (ej. Tarjeta)

## Recursos Disponibles

### `resource://gastos`
Devuelve el listado completo de gastos en formato JSON.

## Estructura del Proyecto

- `main.py`: Código principal del servidor MCP.
- `gastos.csv`: Base de datos local (se crea automáticamente).

# 🔄 CruzarCuentas

**Aplicación en Python para rellenar automáticamente los números de cuenta en archivos de facturas a partir de un listado de clientes.**


---

## 🧩 Características

- Interfaz gráfica moderna
- Relleno automático de cuentas
- Detección de errores por similitud
- Colores según coincidencias:
  - ✅ Exactas
  - 🟧 Dudosas
  - 🟥 No encontradas
- Compatible con `.xlsx`
- Botón para abrir archivo generado

---

## 🖥️ Requisitos

- Python 3.8 o superior
- Librerías:
  - `pandas`
  - `openpyxl`
  - `difflib`

Instálalas con:

```bash
pip install -r requirements.txt
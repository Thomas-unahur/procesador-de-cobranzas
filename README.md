# Auditoría de Cobranzas y Cálculo de Mora (Python + Pandas)

Script que limpia un CSV de facturas, detecta registros sospechosos y genera un reporte en Excel con la deuda vencida por cliente y su antigüedad de mora.

> **Nota:** proyecto de práctica desarrollado con asistencia de IA. Los datos incluidos (`facturas_prueba.csv`) son ficticios.

## Qué problema resuelve

Los registros de cobranzas suelen llegar con formatos mezclados (montos como `$ 185.000,00` y `125000.50`, fechas en distintos formatos, nombres de clientes escritos de varias maneras) y con duplicados o errores de carga. Revisar todo eso a mano en Excel es lento y propenso a errores. Este script automatiza la limpieza, separa lo dudoso para que una persona lo revise y calcula la mora.

## Qué hace

1. **Limpieza y normalización**
   * Unifica nombres de clientes: espacios repetidos, capitalización y algunas siglas (`S.A.`, `SRL`, `GBA`, `BA`).
   * Convierte importes en formato regional (`$ 185.000,00`) o internacional (`125000.50`) a número.
   * Interpreta fechas en `YYYY-MM-DD`, `DD/MM/YYYY` y `YYYY/MM/DD`.
   * Normaliza el estado de pago a `Pagado` o `Pendiente`.
2. **Auditoría de datos** (los registros dudosos se separan y no entran al cálculo de deuda)
   * Facturas con ID duplicado.
   * Posibles duplicados de negocio: mismo cliente, monto y vencimiento con distinto ID.
   * Montos faltantes, negativos, en cero o menores a $1.000.
   * Fechas de vencimiento faltantes o ilegibles.
   * Montos extremos: mayores a Q3 + 3 × IQR, calculado sobre los montos superiores a $1.000.
3. **Cálculo de deuda y reporte (`.xlsx`)**
   * Calcula días de mora contra la fecha de corte y clasifica en tramos: al día, 1 a 30, 31 a 60 y más de 60 días.
   * Genera un libro con tres hojas: **Resumen Deuda Clientes** (indicadores, tabla y gráfico), **Facturas Limpias** y **Auditoria Anomalias**.

## Cómo ejecutarlo

```bash
pip install pandas numpy openpyxl
python procesar_cobranzas.py
```

El script lee `facturas_prueba.csv` (en la misma carpeta) y genera `reporte_cobranzas.xlsx`.

El CSV debe tener las columnas: `id_factura`, `cliente`, `monto`, `fecha_vencimiento`, `estado_pago`.

## Limitaciones conocidas

Cosas que sé que todavía no resuelve y que mejoraría:

* **Montos con un solo separador:** valores como `1.500` o `1,500` se interpretan como 1,5 (sin distinguir miles de decimales).
* **Umbrales fijos:** el mínimo de $1.000 y el factor 3 del criterio IQR están fijos en el código y habría que ajustarlos según el negocio.
* **Montos negativos:** se envían a auditoría; no se distinguen las notas de crédito válidas.
* **Siglas societarias:** solo se corrigen algunas (`S.A.`, `SRL`, etc.).

## Tecnologías

Python 3, Pandas, NumPy, openpyxl.

import pandas as pd
import json
from datetime import datetime
from num2words import num2words
from dateutil import parser

# 🔹 Leer el archivo Excel
df = pd.read_excel("FACTURAR.xlsx")

# 🔹 Reemplazar NaN por vacío para algunos campos opcionales
df = df.fillna("")

# 🔹 Función auxiliar para fechas
def fmt_fecha(valor, formato="%d/%m/%Y", incluir_hora=False):
    if pd.isna(valor):
        return None
    if isinstance(valor, str):
        try:
            dt = parser.parse(valor)
        except Exception:
            return valor
    else:
        dt = valor
    if incluir_hora:
        return dt.strftime("%I:%M:%S %p").lower()  # Ejemplo: 01:19:00 pm
    return dt.strftime(formato)

# 🔹 Función para convertir montos a letras en español
# def monto_a_letras(monto):
#     """
#     Convierte un monto numérico a su representación en letras,
#     manejando enteros y decimales.
    
#     Ejemplo:
#     monto_a_letras(3969.9) -> "tres mil novecientos sesenta y nueve punto nueve"
#     """
#     try:
#         monto_str = str(monto)
#         if '.' in monto_str:
#             partes = monto_str.split('.')
#             entero = int(partes[0])
#             decimal = int(partes[1])
            
#             entero_letras = num2words(entero, lang='es')
#             decimal_letras = num2words(decimal, lang='es')
            
#             return f"{entero_letras} punto {decimal_letras}"
#         else:
#             return num2words(int(monto), lang='es')
#     except Exception as e:
#         print(f"Error: {e}")
#         return None


def monto_a_letras_b(monto):
    """
    Convierte un monto numérico a su representación en letras,
    con la palabra 'punto' para los decimales, sin añadir la moneda.

    Ejemplo:
    monto_a_letras(3969.9) -> "tres mil novecientos sesenta y nueve punto nueve"
    """
    try:
        monto_str = str(monto)
        # Verifica si el número tiene una parte decimal
        if '.' in monto_str:
            # Divide el número en partes entera y decimal
            partes = monto_str.split('.')
            entero = int(partes[0])
            decimal = int(partes[1])

            # Convierte ambas partes a letras
            entero_letras = num2words(entero, lang='es')
            decimal_letras = num2words(decimal, lang='es')

            # Combina las partes con "punto"
            return f"{entero_letras} punto {decimal_letras}"
        else:
            # Si no hay decimales, solo convierte la parte entera
            return num2words(int(monto), lang='es')
            
    except (ValueError, IndexError):
        # Maneja errores en el formato de entrada
        return None


def monto_a_letras(monto, moneda="dolares"):
    """
    Convierte un monto numérico a su representación en letras,
    usando "centavos" o "céntimos" para los decimales.

    Parámetros:
    monto (float): El valor numérico a convertir.
    moneda (str): 'dolares' o 'bolivares'. Por defecto es 'dolares'.

    Retorna:
    str: El monto en letras.
    """
    try:
        monto = float(monto)
        entero = int(monto)
        decimales = int(round((monto - entero) * 100))
        
        entero_letras = num2words(entero, lang='es')

        if moneda.lower() == "bolivares":
            if decimales > 0:
                decimales_letras = num2words(decimales, lang='es')
                return f"{entero_letras} con {decimales_letras} céntimos"
            else:
                return f"{entero_letras} bolívares"
        
        elif moneda.lower() == "dolares":
            if decimales > 0:
                decimales_letras = num2words(decimales, lang='es')
                return f"{entero_letras} con {decimales_letras} centavos"
            else:
                return f"{entero_letras}" # No se añade la palabra "dólares"
        
        else:
            return "Moneda no reconocida."

    except (ValueError, TypeError):
        return None


# 🔹 Función para transformar una fila en JSON con redondeo a 2 decimales y None visibles
def transformar_fila(row):
    fecha_emision = fmt_fecha(row["Fecha Emision"])  # Solo fecha
    fecha_vencimiento = fmt_fecha(row["Fecha Vencimiento"])
    hora_emision = fmt_fecha(row["Fecha Emision"], incluir_hora=True) if row["Fecha Emision"] else None
    fecha_pago = fmt_fecha(row["Fecha Pago"])

    # 🔹 Redondear montos a 2 decimales
    monto_bs = round(float(row["bolivares"]), 2)
    monto_usd = round(float(row["Total"]), 0)
    bolivares_sin_iva = round(float(row["bolivares sin iva"]), 2)
    precio_sin_iva = round(float(row["precio sin iva"]), 2)
    iva_bs = round(monto_bs - bolivares_sin_iva, 2)
    iva_usd = round(monto_usd - precio_sin_iva, 2)
    pro=m = round(float(row["promedio"]), 0)

    return {
        "DocumentoElectronico": {
            "Encabezado": {
                "IdentificacionDocumento": {
                    "TipoDocumento": "01",
                    "NumeroDocumento": str(row["Correlativo"]),
                    "TipoProveedor": None,
                    "TipoTransaccion": None,
                    "NumeroPlanillaImportacion": None,
                    "NumeroExpedienteImportacion": None,
                    "SerieFacturaAfectada": None,
                    "NumeroFacturaAfectada": None,
                    "FechaFacturaAfectada": None,
                    "MontoFacturaAfectada": None,
                    "ComentarioFacturaAfectada": None,
                    "RegimenEspTributacion": None,
                    "FechaEmision": fecha_emision,
                    "FechaVencimiento": fecha_vencimiento,
                    "HoraEmision": hora_emision,
                    "Anulado": False,
                    "TipoDePago": "Inmediato",
                    "Serie": "",
                    "Sucursal": "002",
                    "TipoDeVenta": "Interna",
                    "Moneda": "BSD",
                    "TransaccionId": f"FA{row['Correlativo']}",
                    "UrlPdf": None
                },
                "Vendedor": None,
                "Comprador": {
                    "TipoIdentificacion": "V",
                    "NumeroIdentificacion": str(row["DNI/C.I./C.C./IFE"]),
                    "RazonSocial": row["Cliente"],
                    "Direccion": row["Dirección"],
                    "Ubigeo": None,
                    "Pais": "VE",
                    "Notificar": None,
                    "Telefono": [str(row["Telefono"])],
                    "Correo": [row["Correo"]],
                    "OtrosEnvios": None
                },
                "SujetoRetenido": None,
                "Tercero": None,
                "Totales": {
                    "NroItems": "1",
                    "MontoGravadoTotal": str(monto_bs),
                    "MontoExentoTotal": "0.00",
                    "MontoPercibidoTotal": "0.00",
                    "SubtotalAntesDescuento": str(bolivares_sin_iva),
                    "TotalDescuento": None,
                    "TotalRecargos": None,
                    "Subtotal": str(bolivares_sin_iva),
                    "TotalIVA": str(iva_bs),
                    "MontoTotalConIVA": str(monto_bs),
                    "TotalAPagar": str(monto_bs),
                    "MontoEnLetras":  monto_a_letras_b(monto_bs),
                    "ListaRecargo": None,
                    "ListaDescBonificacion": None,
                    "ImpuestosSubtotal": [
                        {
                            "CodigoTotalImp": "G",
                            "AlicuotaImp": "16.00",
                            "BaseImponibleImp": str(bolivares_sin_iva),
                            "ValorTotalImp": str(iva_bs)
                        }
                    ],
                    "FormasPago": [
                        {
                            "Descripcion": row["Forma de Pago"],
                            "Fecha": fecha_pago,
                            "Forma": "01",
                            "Monto": str(monto_bs),
                            "Moneda": "BSD",
                            "TipoCambio": "0.00"
                        }
                    ],
                    "TotalIGTF": None,
                    "TotalIGTF_VES": None
                },
                "TotalesRetencion": None,
                "TotalesOtraMoneda": {
                    "Moneda": "USD",
                    "TipoCambio": str(row["Tasa"]),
                    "MontoGravadoTotal": str(precio_sin_iva),
                    "MontoPercibidoTotal": "0.00",
                    "MontoExentoTotal": "0.00",
                    "Subtotal": str(precio_sin_iva),
                    "TotalAPagar": str(monto_usd),
                    "TotalIVA": str(iva_usd),
                    "MontoTotalConIVA": str(monto_usd),
                    "MontoEnLetras": monto_a_letras(monto_usd),
                    "SubtotalAntesDescuento": str(precio_sin_iva),
                    "TotalDescuento": None,
                    "TotalRecargos": None,
                    "ListaRecargo": None,
                    "ListaDescBonificacion": None,
                    "ImpuestosSubtotal": [
                        {
                            "CodigoTotalImp": "G",
                            "AlicuotaImp": "16.00",
                            "BaseImponibleImp": str(precio_sin_iva),
                            "ValorTotalImp": str(iva_usd)
                        }
                    ]
                },
                "Orden": None
            },
            "DetallesItems": [
                {
                    "NumeroLinea": "1",
                    "CodigoCIIU": None,
                    "CodigoPLU": "005",
                    "IndicadorBienoServicio": "2",
                    "Descripcion": row["Plan"],
                    "Cantidad": "1",
                    "UnidadMedida": "4L",
                    "PrecioUnitario": str(bolivares_sin_iva),
                    "PrecioUnitarioDescuento": None,
                    "MontoBonificacion": None,
                    "DescripcionBonificacion": None,
                    "DescuentoMonto": "0.00",
                    "RecargoMonto": "0",
                    "PrecioItem": str(bolivares_sin_iva),
                    "PrecioAntesDescuento":str(bolivares_sin_iva),
                    "CodigoImpuesto": "G",
                    "TasaIVA": "16",
                    "ValorIVA": str(iva_bs),
                    "ValorTotalItem":str(bolivares_sin_iva),
                    "InfoAdicionalItem": [],
                    "ListaItemOTI": None
                }
            ],
            "DetallesRetencion": None,
            "Viajes": None,
            "InfoAdicional": [
                {"Campo": "Contrato", "Valor": str(row["ID Servicio"])},
                {"Campo": "Mes1", "Valor": "JUL"},
                {"Campo": "Mes2", "Valor": "AGO"},
                {"Campo": "Mes3", "Valor": "SEP"},
                {"Campo": "Mes4", "Valor": "OCT"},
                {"Campo": "Mes5", "Valor": "NOV"},
                {"Campo": "Mes6", "Valor": "DIC"},
                {"Campo": "Cmes1", "Valor": str(row["navegacion"])},
                {"Campo": "Cmes2", "Valor":  str(row["navegacion2"])},
                {"Campo": "Cmes3", "Valor": "0"},
                {"Campo": "Cmes4", "Valor": "0"},
                {"Campo": "Cmes5", "Valor": "0"},
                {"Campo": "Cmes6", "Valor": "0"},
                {"Campo": "Promedio", "Valor": str(int(row["promedio"]))},
            ],
            "GuiaDespacho": None,
            "Transporte": None,
            "EsLote": True,
            "EsMinimo": None
        }
    }

# 🔹 Generar un archivo JSON por fila
for _, row in df.iterrows():
    data = transformar_fila(row)
    correlativo = str(row["Correlativo"]).zfill(6)
    filename = f"0{correlativo}.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"✅ Archivo generado: {filename}")

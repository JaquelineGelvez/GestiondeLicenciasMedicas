import pandas as pd
from datetime import datetime

# Definición del archivo de la base de datos Excel
excel_file = "Base de Datos .xlsx"

# =====================================================================
# BLOQUE 1: CARGA DE LA BASE DE DATOS
# =====================================================================
# Intentamos leer el archivo de Excel y cargar las hojas necesarias.
# - 'Hoja 1' contiene los datos de los empleados (legajo, nombre,
#   saldo de días, estado).
# - 'Hoja 2' contiene el historial de solicitudes de licencias
#   registradas.
try:
    empleados = pd.read_excel(excel_file, sheet_name="Hoja 1")
    solicitudes = pd.read_excel(excel_file, sheet_name="Hoja 2")
except Exception as e:
    print(f"Error al cargar la base de datos: {e}")
    exit()

print("===================================")
print(" SISTEMA DE LICENCIAS MÉDICAS ")
print("===================================")

# =====================================================================
# BLOQUE 2: SOLICITUD Y VALIDACIÓN DEL LEGAJO (MÁXIMO 3 INTENTOS)
# =====================================================================
# Se le solicita al usuario su número de legajo por consola.
# Si ingresa un valor no numérico o un legajo que no existe en la
# base de datos, el sistema le permite un máximo de 3 intentos.
# Al 3er intento fallido, muestra un mensaje de despedida
# y finaliza la ejecución.
intentos_legajo = 0
while True:
    legajo_str = input("Ingrese su número de legajo: ").strip()
    try:
        legajo = int(legajo_str)
        # Buscar al empleado por su Nº de legajo en 'Hoja 1'
        empleado = empleados[empleados["Nº de legajo"] == legajo]
        if empleado.empty:
            intentos_legajo += 1
            if intentos_legajo >= 3:
                print("\nNúmero de legajo incorrecto.")
                print("Gracias por utilizar nuestros servicios.")
                exit()
            print("Número de legajo incorrecto. Intente nuevamente.\n")
        else:
            break
    except ValueError:
        intentos_legajo += 1
        if intentos_legajo >= 3:
            print("\nNúmero de legajo incorrecto.")
            print("Gracias por utilizar nuestros servicios.")
            exit()
        print("Número de legajo incorrecto. Intente nuevamente.\n")

# Extraemos los datos del empleado validado
nombre = empleado.iloc[0]["Nombre"]
apellido = empleado.iloc[0]["Apellido"]
estado_empleado = empleado.iloc[0]["Estado del empleado"]

# =====================================================================
# BLOQUE 3: VALIDACIÓN DEL ESTADO DEL EMPLEADO (ACTIVO)
# =====================================================================
# Verificamos si el empleado se encuentra "activo".
# Si su estado es "inactivo", se rechaza su solicitud por inactividad
# y se le solicita comunicarse con Recursos Humanos a la brevedad.
# Si tiene cualquier otro estado no activo, se deniega de igual forma.
if str(estado_empleado).strip().lower() != "activo":
    if str(estado_empleado).strip().lower() == "inactivo":
        print(
            f"\nLo sentimos, {nombre} {apellido}.\n"
            "Su solicitud de licencia ha sido rechazada debido a su "
            "inactividad en la empresa.\n"
            "Por favor, comuníquese con Recursos Humanos a la brevedad."
        )
    else:
        print(
            f"\nLo sentimos, {nombre} {apellido}. "
            f"Su estado actual es '{estado_empleado}' "
            "y no puede solicitar licencias."
        )
    print("Fin del proceso.")
    exit()

print(f"\nBienvenido/a {nombre} {apellido}")

# =====================================================================
# BLOQUE 4: VALIDACIÓN DEL CERTIFICADO MÉDICO (MÁXIMO 3 INTENTOS)
# =====================================================================
# Se consulta al usuario si adjuntó el certificado
# correspondiente (si/no).
# Si responde "no", el proceso se cancela, ya que el certificado
# es obligatorio. Si ingresa una respuesta inválida (distinta de
# 'si' o 'no'), tiene hasta 3 intentos antes de que el programa
# se cierre por error de entrada.
intentos_cert = 0
while True:
    certificado = (
        input("\n¿Adjuntó certificado médico? (si/no): ")
        .strip()
        .lower()
    )
    if certificado == "si":
        break
    elif certificado == "no":
        print("\nDebe adjuntar certificado médico.")
        reintento = (
            input("¿Desea volver a intentarlo? (si/no): ")
            .strip()
            .lower()
        )
        if reintento != "si":
            print("Solicitud cancelada.")
            exit()
    else:
        intentos_cert += 1
        if intentos_cert >= 3:
            print("\nTerminó la operación, vuelva a intentarlo más tarde.")
            exit()
        print("Por favor, responda 'si' o 'no'.")

# =====================================================================
# BLOQUE 5: INGRESO DEL DIAGNÓSTICO MÉDICO (MÁXIMO 3 INTENTOS)
# =====================================================================
# Se solicita que describa el diagnóstico médico.
# Si se deja vacío, el sistema avisa y le da hasta 3 intentos
# antes de salir.
diagnostico = input("\nIngrese diagnóstico médico: ").strip()
intentos_diag = 1
while not diagnostico:
    if intentos_diag >= 3:
        print("\nTerminó la operación, vuelva a intentarlo más tarde.")
        exit()
    print("El diagnóstico no puede estar vacío.")
    diagnostico = input("Ingrese diagnóstico médico: ").strip()
    intentos_diag += 1

# =====================================================================
# BLOQUE 6: ASIGNACIÓN AUTOMÁTICA DE LA FECHA DE INICIO
# =====================================================================
# Para cumplir con el diagrama BPMN del usuario, el chatbot no
# solicita la fecha de inicio por consola, sino que la asigna
# automáticamente como el día de hoy.
fecha_inicio = datetime.today()

# =====================================================================
# BLOQUE 7: SOLICITUD Y VALIDACIÓN DE DÍAS SOLICITADOS (MÁXIMO 3 INTENTOS)
# =====================================================================
# Se le pide al usuario la cantidad de días que requiere para su
# licencia. Se valida que el valor ingresado sea un número entero
# positivo. Si ingresa valores inválidos (letras o números menores o
# iguales a 0), el sistema le permite un máximo de 3 intentos antes
# de abortar la operación.
dias_disponibles = empleado.iloc[0]["Dias de licencia"]
intentos_dias = 0
while True:
    dias_str = input(
        f"Ingrese cantidad de días solicitados "
        f"(Disponibles: {dias_disponibles}): "
    ).strip()
    try:
        dias = int(dias_str)
        if dias <= 0:
            print("La cantidad de días debe ser un número entero positivo.")
            intentos_dias += 1
            if intentos_dias >= 3:
                print("\nTerminó la operación, vuelva a intentarlo más tarde.")
                exit()
            continue
        break
    except ValueError:
        print("Cantidad de días inválida. Ingrese un número entero.")
        intentos_dias += 1
        if intentos_dias >= 3:
            print("\nTerminó la operación, vuelva a intentarlo más tarde.")
            exit()

# =====================================================================
# BLOQUE 8: EVALUACIÓN DE LA SOLICITUD Y DESCUENTO DE DÍAS
# =====================================================================
# Se compara la cantidad de días solicitados contra los días
# disponibles.
# - Si tiene saldo suficiente: la solicitud se pre-aprueba
#   ('aprobado') y se le restan los días a su saldo disponible
#   en la 'Hoja 1' del DataFrame.
# - Si no tiene saldo: la solicitud se registra como 'rechazado'
#   en 'Hoja 2', y no se modifican los días del empleado en 'Hoja 1'.
if dias <= dias_disponibles:
    estado_solicitud = "aprobado"
    observaciones = "Correcto"
    # Descontar días disponibles del empleado en la base de datos
    empleados.loc[
        empleados["Nº de legajo"] == legajo, "Dias de licencia"
    ] -= dias
    solicitud_ok = True
else:
    estado_solicitud = "rechazado"
    observaciones = "Supera días disponibles"
    solicitud_ok = False

# =====================================================================
# BLOQUE 9: REGISTRO DE LA SOLICITUD EN LA BASE DE DATOS
# =====================================================================
# Se añade el nuevo registro a la 'Hoja 2' con los datos provistos:
# Legajo, Fecha de solicitud (nombre), Fecha de inicio, Días,
# Diagnóstico, Certificado (Si), Estado de solicitud y Observaciones.
print("\nRegistrando solicitud...")

nueva_solicitud = pd.DataFrame([{
    "Legajo": legajo,
    # Guarda el nombre para mantener la convención del Excel
    "Fecha de solicitud": nombre,
    "Fecha de inicio": fecha_inicio,
    "Dias solicitados": dias,
    "Diagnostico": diagnostico,
    "certificado": "Si",
    "estado/solicitud": estado_solicitud,
    "observaciones": observaciones
}])

# Concatenamos la nueva fila al histórico de solicitudes
solicitudes = pd.concat([solicitudes, nueva_solicitud], ignore_index=True)

# Guardamos los cambios de vuelta en el archivo Excel,
# sobreescribiendo ambas hojas para actualizar tanto el saldo de días
# como el registro de solicitudes.
try:
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="w") as writer:
        empleados.to_excel(writer, sheet_name="Hoja 1", index=False)
        solicitudes.to_excel(writer, sheet_name="Hoja 2", index=False)
except Exception as e:
    print(f"Error al guardar los datos en el archivo Excel: {e}")
    exit()

# =====================================================================
# BLOQUE 10: SALIDA EN PANTALLA Y MENSAJES FINALES
# =====================================================================
# Imprimimos en pantalla el resumen de la operación según corresponda.
print("\n===================================")
if solicitud_ok:
    print("SOLICITUD PROCESADA")
else:
    print("SOLICITUD RECHAZADA")
print("===================================")

print(f"Empleado: {nombre} {apellido}")
print(f"Fecha de inicio: {fecha_inicio.strftime('%d/%m/%Y')}")
print(f"Diagnóstico: {diagnostico}")
print(f"Días solicitados: {dias}")
print(f"Estado: {estado_solicitud.capitalize()}")

# Mensajes informativos dinámicos de acuerdo al estado
if solicitud_ok:
    print(
        "\nSi bien la solicitud ha sido aprobada y registrada correctamente, "
        "tenga en cuenta que toda la información y la documentación "
        "presentada quedan sujetas a revisión final por parte de nuestro "
        "personal."
    )
else:
    print(
        "\nLa solicitud ha sido rechazada debido a que excede "
        "el límite de días de licencia que tiene disponibles actualmente."
    )

print("\nFin del proceso.")

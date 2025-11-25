# --- Dropdown Options ---

# --- Data untuk auto-fill ---
EQUIPO_AUTOFILL_DATA = {
    "ALMEMO": {
        "marca": "MA710",
        "tipo": "Registrador de Temperatura",
        "fecha": "05/11/2014"  # Contoh tanggal, silakan diubah
    },
    "TERMOHIGRÓMETRO": {
        "marca": "MA24702S",
        "tipo": "Medición Temperatura Ambiente",
        "fecha": "31/10/2024"  # Contoh tanggal, silakan diubah
    },
    "CAMARA ENDURANCIA": {
        "marca": "CET10/15312",
        "tipo": "Dycometal" 
    },
    "SONDA TIPO T": {
        "fecha": "23/09/2025"  # Contoh tanggal, silakan diubah
    },
}

# TODO: Add more equipment options here
EQUIPMENT_OPTIONS = [
    "ALMEMO",
    "TERMOHIGRÓMETRO",
    "CAMARA ENDURANCIA",
    "SONDA TIPO T",
]

# TODO: Add more brand/model options here
MARCA_OPTIONS = [
    "MA710",
    "MA24702S",
    "CET10/15312",
    "2024103000",
    "2024103001",
    "2024103002",
    "2024103003",
    "2024103004",
    "2024103005",
    "2024103006",
    "2024103007",
    "2024103008",
    "2024103009",
    "2024103010",
    "2024103011",
]

# TODO: Add more type/application options here
TIPO_OPTIONS = [
    "Registrador de Temperatura",
    "Medición Temperatura Ambiente",
    "Dycometal",
    "Carcasa int.",
    "Carcasa ext.",
    "Disipador",
    "Difusor",
    "PCB",
    "Driver",
    "Tc LED",
    "T. Ambiente",
    "Reflector",
    "Tc LED Calido",
    "Tc LED Frio",
    "Lente",
]

# Temperature measurement point options
PUNTO_OPTIONS = [
    "Carcasa int.",
    "Carcasa ext.",
    "Disipador",
    "Difusor",
    "PCB",
    "Driver",
    "Tc LED",
    "T. Ambiente",
    "Reflector",
    "Tc LED Calido",
    "Tc LED Frio",
    "Lente"
]

# Temperature unit options
UNIDAD_OPTIONS = [
    "°C",
    "°F",
    "K"
]

# Common temperature limits
LIMITE_OPTIONS = [
   "50", "55", "60", "70", "75", "80", "85", "90", "95", "100", "105", "115", "120", "125", "130", "135", "140", "165", "170", "200", "210", "225", "230", "250","N/A"
]

# Result options
RESULT_OPTIONS = [
    "Pass",
    "Fail",
    "N/A"
]

# Title options for montage photos
TITLE_OPTIONS = [
   "Montaje Final 1",
    "Montaje Final 2",
    "Montaje Final 3",
    "Montaje Final 4",
    "Montaje Final 5",
    "Montaje Final 6",
    "Montaje Final 7",
    "Termohigrometro Temperatura",
    "Termohigrometro Humedad",
    "Termohigrometro Presión Atmosférica",
]

# Ensayo type options
ENSAYO_OPTIONS = [
    "funcionamiento normal",
    "funcionamiento anormal",
]

# Opsi Código berdasarkan Equipo
CODIGO_OPTIONS_BY_EQUIPO = {
    "ALMEMO": ["ALM-001", "ALM-002", "ALM-003"],
    "TERMOHIGRÓMETRO": ["TH-101", "TH-102"],
    "CAMARA ENDURANCIA": ["CE-201", "CE-202"],
    "SONDA TIPO T": [
        "STT-A-01", "STT-A-02", "STT-A-03", "STT-A-04",
        "STT-B-01", "STT-B-02", "STT-B-03", "STT-B-04",
        "STT-C-01", "STT-C-02", "STT-C-03", "STT-C-04"
    ],
    # Tambahkan equipo dan kode lainnya di sini
}

# Gabungkan semua kode untuk inisialisasi dropdown awal
ALL_CODIGO_OPTIONS = sorted(list(set(code for codes in CODIGO_OPTIONS_BY_EQUIPO.values() for code in codes)))

# --- Definisi Placeholders & Input Fields ---
FIELD_DEFINITIONS = {
    # Header fields
    "NO_TEST": {"placeholder": "[NO_TEST]", "label": "Nº de Test Plan", "type": "text"},
    "REV": {"placeholder": "[REV]", "label": "Revisión", "type": "text"},
    "DATE": {"placeholder": "[DATE]", "label": "Fecha de emisión (DD/MM/YYYY)", "type": "date"},

    # 0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO
    "TEXT1": {"placeholder": "[TEXT1]", "label": "Solicitante", "type": "text"},
    "TEXT4": {"placeholder": "[TEXT4]", "label": "Operario del ensayo", "type": "text"},
    "TEXT2": {"placeholder": "[TEXT2]", "label": "Departamento", "type": "text"},
    "TEXT5": {"placeholder": "[TEXT5]", "label": "Responsable del ensayo", "type": "text"},
    "TEXT3": {"placeholder": "[TEXT3]", "label": "Fecha de solicitud (DD/MM/YYYY)", "type": "date"},

    # 1. INFORMACIÓN GENERAL DEL PRODUCTO
    "TEXT6": {"placeholder": "[TEXT6]", "label": "Referencia del modelo ensayado", "type": "text"},
    "TEXT12": {"placeholder": "[TEXT12]", "label": "Familia", "type": "text"},
    "TEXT7": {"placeholder": "[TEXT7]", "label": "Aplicación", "type": "text"},
    "TEXT8": {"placeholder": "[TEXT8]", "label": "Fuente de luz", "type": "text"},
    "TEXT13": {"placeholder": "[TEXT13]", "label": "Fuente de alimentación (driver) ", "type": "text"},

    # 1.1. CONDICIONES DEL ENSAYO
    "TEXT9": {"placeholder": "[TEXT9]", "label": "Ensayo térmico realizado en", "type": "dropdown", "options": ENSAYO_OPTIONS},
    "TEXT10": {"placeholder": "[TEXT10]", "label": "Temperatura de color ensayada (CCT)", "type": "text"},
    "TEXT11": {"placeholder": "[TEXT11]", "label": "Luminaria alimentada a", "type": "text"},

    # 2. EQUIPOS Y MÉTODOS UTILIZADOS (up to 12)
    "EQUIPO1": {"placeholder": "[EQUIPO1]", "label": "Equipo 1", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA1": {"placeholder": "[MARCA1]", "label": "Marca/Modelo 1", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO1": {"placeholder": "[TIPO1]", "label": "Tipo/Aplicación 1", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA1": {"placeholder": "[FECHA1]", "label": "Fecha de calibración 1", "type": "date"},
    "OBSER1": {"placeholder": "[OBSER1]", "label": "Código 1", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO2": {"placeholder": "[EQUIPO2]", "label": "Equipo 2", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA2": {"placeholder": "[MARCA2]", "label": "Marca/Modelo 2", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO2": {"placeholder": "[TIPO2]", "label": "Tipo/Aplicación 2", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA2": {"placeholder": "[FECHA2]", "label": "Fecha de calibración 2", "type": "date"},
    "OBSER2": {"placeholder": "[OBSER2]", "label": "Código 2", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO3": {"placeholder": "[EQUIPO3]", "label": "Equipo 3", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA3": {"placeholder": "[MARCA3]", "label": "Marca/Modelo 3", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO3": {"placeholder": "[TIPO3]", "label": "Tipo/Aplicación 3", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA3": {"placeholder": "[FECHA3]", "label": "Fecha de calibración 3", "type": "date"},
    "OBSER3": {"placeholder": "[OBSER3]", "label": "Código 3", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO4": {"placeholder": "[EQUIPO4]", "label": "Equipo 4", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA4": {"placeholder": "[MARCA4]", "label": "Marca/Modelo 4", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO4": {"placeholder": "[TIPO4]", "label": "Tipo/Aplicación 4", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA4": {"placeholder": "[FECHA4]", "label": "Fecha de calibración 4", "type": "date"},
    "OBSER4": {"placeholder": "[OBSER4]", "label": "Código 4", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO5": {"placeholder": "[EQUIPO5]", "label": "Equipo 5", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA5": {"placeholder": "[MARCA5]", "label": "Marca/Modelo 5", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO5": {"placeholder": "[TIPO5]", "label": "Tipo/Aplicación 5", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA5": {"placeholder": "[FECHA5]", "label": "Fecha de calibración 5", "type": "date"},
    "OBSER5": {"placeholder": "[OBSER5]", "label": "Código 5", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO6": {"placeholder": "[EQUIPO6]", "label": "Equipo 6", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA6": {"placeholder": "[MARCA6]", "label": "Marca/Modelo 6", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO6": {"placeholder": "[TIPO6]", "label": "Tipo/Aplicación 6", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA6": {"placeholder": "[FECHA6]", "label": "Fecha de calibración 6", "type": "date"},
    "OBSER6": {"placeholder": "[OBSER6]", "label": "Código 6", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO7": {"placeholder": "[EQUIPO7]", "label": "Equipo 7", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA7": {"placeholder": "[MARCA7]", "label": "Marca/Modelo 7", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO7": {"placeholder": "[TIPO7]", "label": "Tipo/Aplicación 7", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA7": {"placeholder": "[FECHA7]", "label": "Fecha de calibración 7", "type": "date"},
    "OBSER7": {"placeholder": "[OBSER7]", "label": "Código 7", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO8": {"placeholder": "[EQUIPO8]", "label": "Equipo 8", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA8": {"placeholder": "[MARCA8]", "label": "Marca/Modelo 8", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO8": {"placeholder": "[TIPO8]", "label": "Tipo/Aplicación 8", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA8": {"placeholder": "[FECHA8]", "label": "Fecha de calibración 8", "type": "date"},
    "OBSER8": {"placeholder": "[OBSER8]", "label": "Código 8", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO9": {"placeholder": "[EQUIPO9]", "label": "Equipo 9", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA9": {"placeholder": "[MARCA9]", "label": "Marca/Modelo 9", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO9": {"placeholder": "[TIPO9]", "label": "Tipo/Aplicación 9", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA9": {"placeholder": "[FECHA9]", "label": "Fecha de calibración 9", "type": "date"},
    "OBSER9": {"placeholder": "[OBSER9]", "label": "Código 9", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO10": {"placeholder": "[EQUIPO10]", "label": "Equipo 10", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA10": {"placeholder": "[MARCA10]", "label": "Marca/Modelo 10", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO10": {"placeholder": "[TIPO10]", "label": "Tipo/Aplicación 10", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA10": {"placeholder": "[FECHA10]", "label": "Fecha de calibración 10", "type": "date"},
    "OBSER10": {"placeholder": "[OBSER10]", "label": "Código 10", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO11": {"placeholder": "[EQUIPO11]", "label": "Equipo 11", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA11": {"placeholder": "[MARCA11]", "label": "Marca/Modelo 11", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO11": {"placeholder": "[TIPO11]", "label": "Tipo/Aplicación 11", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA11": {"placeholder": "[FECHA11]", "label": "Fecha de calibración 11", "type": "date"},
    "OBSER11": {"placeholder": "[OBSER11]", "label": "Código 11", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},
    "EQUIPO12": {"placeholder": "[EQUIPO12]", "label": "Equipo 12", "type": "dropdown", "options": EQUIPMENT_OPTIONS},
    "MARCA12": {"placeholder": "[MARCA12]", "label": "Marca/Modelo 12", "type": "dropdown", "options": MARCA_OPTIONS},
    "TIPO12": {"placeholder": "[TIPO12]", "label": "Tipo/Aplicación 12", "type": "dropdown", "options": TIPO_OPTIONS},
    "FECHA12": {"placeholder": "[FECHA12]", "label": "Fecha de calibración 12", "type": "date"},
    "OBSER12": {"placeholder": "[OBSER12]", "label": "Código 12", "type": "dropdown", "options": ALL_CODIGO_OPTIONS},

    # 3. TEMPERATURAS REGISTRADAS (up to 10)
    "PUNTO1": {"placeholder": "[PUNTO1]", "label": "Punto de Medición 1", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE1": {"placeholder": "[LIMITE1]", "label": "Límite Máximo 1", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP1": {"placeholder": "[TEMP1]", "label": "Temperatura Medida 1", "type": "text"},
    "PUNTO2": {"placeholder": "[PUNTO2]", "label": "Punto de Medición 2", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE2": {"placeholder": "[LIMITE2]", "label": "Límite Máximo 2", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP2": {"placeholder": "[TEMP2]", "label": "Temperatura Medida 2", "type": "text"},
    "PUNTO3": {"placeholder": "[PUNTO3]", "label": "Punto de Medición 3", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE3": {"placeholder": "[LIMITE3]", "label": "Límite Máximo 3", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP3": {"placeholder": "[TEMP3]", "label": "Temperatura Medida 3", "type": "text"},
    "PUNTO4": {"placeholder": "[PUNTO4]", "label": "Punto de Medición 4", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE4": {"placeholder": "[LIMITE4]", "label": "Límite Máximo 4", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP4": {"placeholder": "[TEMP4]", "label": "Temperatura Medida 4", "type": "text"},
    "PUNTO5": {"placeholder": "[PUNTO5]", "label": "Punto de Medición 5", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE5": {"placeholder": "[LIMITE5]", "label": "Límite Máximo 5", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP5": {"placeholder": "[TEMP5]", "label": "Temperatura Medida 5", "type": "text"},
    "PUNTO6": {"placeholder": "[PUNTO6]", "label": "Punto de Medición 6", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE6": {"placeholder": "[LIMITE6]", "label": "Límite Máximo 6", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP6": {"placeholder": "[TEMP6]", "label": "Temperatura Medida 6", "type": "text"},
    "PUNTO7": {"placeholder": "[PUNTO7]", "label": "Punto de Medición 7", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE7": {"placeholder": "[LIMITE7]", "label": "Límite Máximo 7", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP7": {"placeholder": "[TEMP7]", "label": "Temperatura Medida 7", "type": "text"},
    "PUNTO8": {"placeholder": "[PUNTO8]", "label": "Punto de Medición 8", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE8": {"placeholder": "[LIMITE8]", "label": "Límite Máximo 8", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP8": {"placeholder": "[TEMP8]", "label": "Temperatura Medida 8", "type": "text"},
    "PUNTO9": {"placeholder": "[PUNTO9]", "label": "Punto de Medición 9", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE9": {"placeholder": "[LIMITE9]", "label": "Límite Máximo 9", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP9": {"placeholder": "[TEMP9]", "label": "Temperatura Medida 9", "type": "text"},
    "PUNTO10": {"placeholder": "[PUNTO10]", "label": "Punto de Medición 10", "type": "dropdown", "options": PUNTO_OPTIONS},
    "LIMITE10": {"placeholder": "[LIMITE10]", "label": "Límite Máximo 10", "type": "dropdown", "options": LIMITE_OPTIONS},
    "TEMP10": {"placeholder": "[TEMP10]", "label": "Temperatura Medida 10", "type": "text"},
    # Placeholders for units (no UI element)
    "UNIDAD1": {"placeholder": "[UNIDAD1]"}, "UNIDAD2": {"placeholder": "[UNIDAD2]"},
    "UNIDAD3": {"placeholder": "[UNIDAD3]"}, "UNIDAD4": {"placeholder": "[UNIDAD4]"},
    "UNIDAD5": {"placeholder": "[UNIDAD5]"}, "UNIDAD6": {"placeholder": "[UNIDAD6]"},
    "UNIDAD7": {"placeholder": "[UNIDAD7]"}, "UNIDAD8": {"placeholder": "[UNIDAD8]"},
    "UNIDAD9": {"placeholder": "[UNIDAD9]"}, "UNIDAD10": {"placeholder": "[UNIDAD10]"},

    "TEXT14": {"placeholder": "[TEXT14]", "label": "Description", "type": "text"},
    "TEXT15": {"placeholder": "[TEXT15]", "label": "Conclusions", "type": "text"},

    # 3.1. GRÁFICA GENERADA
    "IMAGE1": {"placeholder": "[IMAGE1]", "label": "Imagen 1", "type": "file"},
    "TITLE1": {"placeholder": "[TITLE1]", "label": "Periodo de estabilización", "type": "text"},
    "IMAGE2": {"placeholder": "[IMAGE2]", "label": "Imagen 2", "type": "file"},

    # 4. ESTABILIZACIÓN TÉRMICA (up to 10)

    "VALMIN1": {"placeholder": "[VALMIN1]", "label": "Valor Mínimo 1", "type": "text"},
    "VALMAX1": {"placeholder": "[VALMAX1]", "label": "Valor Máximo 1", "type": "text"},
    "DESVI1": {"placeholder": "[DESVI1]", "label": "Desviación 1", "type": "text"},

    "VALMIN2": {"placeholder": "[VALMIN2]", "label": "Valor Mínimo 2", "type": "text"},
    "VALMAX2": {"placeholder": "[VALMAX2]", "label": "Valor Máximo 2", "type": "text"},
    "DESVI2": {"placeholder": "[DESVI2]", "label": "Desviación 2", "type": "text"},

    "VALMIN3": {"placeholder": "[VALMIN3]", "label": "Valor Mínimo 3", "type": "text"},
    "VALMAX3": {"placeholder": "[VALMAX3]", "label": "Valor Máximo 3", "type": "text"},
    "DESVI3": {"placeholder": "[DESVI3]", "label": "Desviación 3", "type": "text"},

    "VALMIN4": {"placeholder": "[VALMIN4]", "label": "Valor Mínimo 4", "type": "text"},
    "VALMAX4": {"placeholder": "[VALMAX4]", "label": "Valor Máximo 4", "type": "text"},
    "DESVI4": {"placeholder": "[DESVI4]", "label": "Desviación 4", "type": "text"},

    "VALMIN5": {"placeholder": "[VALMIN5]", "label": "Valor Mínimo 5", "type": "text"},
    "VALMAX5": {"placeholder": "[VALMAX5]", "label": "Valor Máximo 5", "type": "text"},
    "DESVI5": {"placeholder": "[DESVI5]", "label": "Desviación 5", "type": "text"},
    
    "VALMIN6": {"placeholder": "[VALMIN6]", "label": "Valor Mínimo 6", "type": "text"},
    "VALMAX6": {"placeholder": "[VALMAX6]", "label": "Valor Máximo 6", "type": "text"},
    "DESVI6": {"placeholder": "[DESVI6]", "label": "Desviación 6", "type": "text"},

    "VALMIN7": {"placeholder": "[VALMIN7]", "label": "Valor Mínimo 7", "type": "text"},
    "VALMAX7": {"placeholder": "[VALMAX7]", "label": "Valor Máximo 7", "type": "text"},
    "DESVI7": {"placeholder": "[DESVI7]", "label": "Desviación 7", "type": "text"},
    
    "VALMIN8": {"placeholder": "[VALMIN8]", "label": "Valor Mínimo 8", "type": "text"},
    "VALMAX8": {"placeholder": "[VALMAX8]", "label": "Valor Máximo 8", "type": "text"},
    "DESVI8": {"placeholder": "[DESVI8]", "label": "Desviación 8", "type": "text"},

    "VALMIN9": {"placeholder": "[VALMIN9]", "label": "Valor Mínimo 9", "type": "text"},
    "VALMAX9": {"placeholder": "[VALMAX9]", "label": "Valor Máximo 9", "type": "text"},
    "DESVI9": {"placeholder": "[DESVI9]", "label": "Desviación 9", "type": "text"},

    "VALMIN10": {"placeholder": "[VALMIN10]", "label": "Valor Mínimo 10", "type": "text"},
    "VALMAX10": {"placeholder": "[VALMAX10]", "label": "Valor Máximo 10", "type": "text"},
    "DESVI10": {"placeholder": "[DESVI10]", "label": "Desviación 10", "type": "text"},
    # Placeholders for stabilization units (no UI element)
    # 5. RESULTADOS (sekarang bagian dari tab Temperatur)
    "RESULT1": {"placeholder": "[RESULT1]", "label": "Resultado 1", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT2": {"placeholder": "[RESULT2]", "label": "Resultado 2", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT3": {"placeholder": "[RESULT3]", "label": "Resultado 3", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT4": {"placeholder": "[RESULT4]", "label": "Resultado 4", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT5": {"placeholder": "[RESULT5]", "label": "Resultado 5", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT6": {"placeholder": "[RESULT6]", "label": "Resultado 6", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT7": {"placeholder": "[RESULT7]", "label": "Resultado 7", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT8": {"placeholder": "[RESULT8]", "label": "Resultado 8", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT9": {"placeholder": "[RESULT9]", "label": "Resultado 9", "type": "dropdown", "options": RESULT_OPTIONS},
    "RESULT10": {"placeholder": "[RESULT10]", "label": "Resultado 10", "type": "dropdown", "options": RESULT_OPTIONS},

    # 7. FOTOGRAFIAS (up to 10)
    "IMAGE3": {"placeholder": "[IMAGE3]", "label": "Imagen 3", "type": "file"},
    "TITLE3": {"placeholder": "[TITLE3]", "label": "Título 3", "type": "text"},
    "IMAGE4": {"placeholder": "[IMAGE4]", "label": "Imagen 4", "type": "file"},
    "TITLE4": {"placeholder": "[TITLE4]", "label": "Título 4", "type": "text"},
    "IMAGE5": {"placeholder": "[IMAGE5]", "label": "Imagen 5", "type": "file"},
    "TITLE5": {"placeholder": "[TITLE5]", "label": "Título 5", "type": "text"},
    "IMAGE6": {"placeholder": "[IMAGE6]", "label": "Imagen 6", "type": "file"},
    "TITLE6": {"placeholder": "[TITLE6]", "label": "Título 6", "type": "text"},
    "IMAGE7": {"placeholder": "[IMAGE7]", "label": "Imagen 7", "type": "file"},
    "TITLE7": {"placeholder": "[TITLE7]", "label": "Título 7", "type": "text"},
    "IMAGE8": {"placeholder": "[IMAGE8]", "label": "Imagen 8", "type": "file"},
    "TITLE8": {"placeholder": "[TITLE8]", "label": "Título 8", "type": "text"},
    "TITLE9": {"placeholder": "[TITLE9]", "label": "Título 9", "type": "text"},
    "IMAGE9": {"placeholder": "[IMAGE9]", "label": "Imagen 9", "type": "file"},
    "TITLE10": {"placeholder": "[TITLE10]", "label": "Título 10", "type": "text"},
    "IMAGE10": {"placeholder": "[IMAGE10]", "label": "Imagen 10", "type": "file"},
    "TITLE11": {"placeholder": "[TITLE11]", "label": "Título 11", "type": "text"},
    "IMAGE11": {"placeholder": "[IMAGE11]", "label": "Imagen 11", "type": "file"},
    "TITLE12": {"placeholder": "[TITLE12]", "label": "Título 12", "type": "text"},
    "IMAGE12": {"placeholder": "[IMAGE12]", "label": "Imagen 12", "type": "file"},

    # 7. FOTO MONTAJE FINAL (up to 10)
    "IMAGE13": {"placeholder": "[IMAGE13]", "label": "Imagen 13", "type": "file"},
    "TITLE13": {"placeholder": "[TITLE13]", "label": "Título 13", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE14": {"placeholder": "[IMAGE14]", "label": "Imagen 14", "type": "file"},
    "TITLE14": {"placeholder": "[TITLE14]", "label": "Título 14", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE15": {"placeholder": "[IMAGE15]", "label": "Imagen 15", "type": "file"},
    "TITLE15": {"placeholder": "[TITLE15]", "label": "Título 15", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE16": {"placeholder": "[IMAGE16]", "label": "Imagen 16", "type": "file"},
    "TITLE16": {"placeholder": "[TITLE16]", "label": "Título 16", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE17": {"placeholder": "[IMAGE17]", "label": "Imagen 17", "type": "file"},
    "TITLE17": {"placeholder": "[TITLE17]", "label": "Título 17", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE18": {"placeholder": "[IMAGE18]", "label": "Imagen 18", "type": "file"},
    "TITLE18": {"placeholder": "[TITLE18]", "label": "Título 18", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE19": {"placeholder": "[IMAGE19]", "label": "Imagen 19", "type": "file"},
    "TITLE19": {"placeholder": "[TITLE19]", "label": "Título 19", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE20": {"placeholder": "[IMAGE20]", "label": "Imagen 20", "type": "file"},
    "TITLE20": {"placeholder": "[TITLE20]", "label": "Título 20", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE21": {"placeholder": "[IMAGE21]", "label": "Imagen 21", "type": "file"},
    "TITLE21": {"placeholder": "[TITLE21]", "label": "Título 21", "type": "dropdown", "options": TITLE_OPTIONS},
    "IMAGE22": {"placeholder": "[IMAGE22]", "label": "Imagen 22", "type": "file"},
    "TITLE22": {"placeholder": "[TITLE22]", "label": "Título 22", "type": "dropdown", "options": TITLE_OPTIONS},
}

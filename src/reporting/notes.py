class NotesOrchestrator:
    def __init__(self, mapping_rules):
        """
        Req 5: Orquesta las tablas de Notas.
        mapping_rules define qué cuentas o categorías entran en qué Nota.
        """
        self.mapping_rules = mapping_rules
        
    def generate_note_tables(self, mapped_tb_df):
        """
        Retorna un diccionario con detalles estructurados validos para tablas Word.
        Ej: {'Nota_10_Inventarios': {'total': 5000, 'detalle': [...]}}
        """
        notes_data = {}
        for note_name, categories in self.mapping_rules.items():
            # Filtrar Trial Balance
            note_df = mapped_tb_df[mapped_tb_df['categoria_financiera'].isin(categories)].copy()
            
            if not note_df.empty:
                # Agrupar para generar la tabla descriptiva de la nota
                table_data = note_df.groupby(['cuenta_id', 'descripcion'])['saldo_final'].sum().reset_index()
                notes_data[note_name] = {
                    "total": table_data['saldo_final'].sum(),
                    "detalle": table_data.to_dict(orient='records')
                }
            else:
                notes_data[note_name] = {"total": 0.0, "detalle": []}
                
        return notes_data


# Registro Maestro de Códigos Únicos de Notas con prefijo '#'
NOTE_REGISTRY = {
    "#N04": {
        "title": "Efectivo y equivalentes al efectivo",
        "sheets": ["Efectivo"],
        "category": "activos_corrientes"
    },
    "#N05": {
        "title": "Otros activos no financieros (Corriente)",
        "sheets": ["Otros activos no financieros, c"],
        "category": "activos_corrientes"
    },
    "#N06": {
        "title": "Deudores comerciales y otras cuentas por cobrar",
        "sheets": ["Deudores"],
        "category": "activos_corrientes"
    },
    "#N07": {
        "title": "Inventarios",
        "sheets": ["Inventarios"],
        "category": "activos_corrientes"
    },
    "#N08": {
        "title": "Activos intangibles distintos de la plusvalía",
        "sheets": ["Intangibles"],
        "category": "activos_no_corrientes"
    },
    "#N09": {
        "title": "Propiedades, planta y equipo",
        "sheets": ["Activo Fijo"],
        "category": "activos_no_corrientes"
    },
    "#N10": {
        "title": "Activos por derechos de uso",
        "sheets": ["Activo por derechos de uso"],
        "category": "activos_no_corrientes"
    },
    "#N11": {
        "title": "Plusvalía",
        "sheets": ["Plusvalia"],
        "category": "activos_no_corrientes"
    },
    "#N12": {
        "title": "Activos y pasivos de impuestos corrientes",
        "sheets": ["Impuestos corrientes"],
        "category": "activos_corrientes"
    },
    "#N13": {
        "title": "Impuestos diferidos",
        "sheets": ["Impuestos Diferidos"],
        "category": "activos_no_corrientes"
    },
    "#N14": {
        "title": "Cuentas por cobrar/pagar a entidades relacionadas",
        "sheets": ["Empresas relacionadas"],
        "category": "pasivos_corrientes"
    },
    "#N15": {
        "title": "Instrumentos / Pasivos financieros",
        "sheets": ["Pasivos financieros "],
        "category": "pasivos_no_corrientes"
    },
    "#N16": {
        "title": "Pasivos por derechos de uso",
        "sheets": ["Pasivos derechos de  uso"],
        "category": "pasivos_no_corrientes"
    },
    "#N17": {
        "title": "Cuentas por pagar comerciales y otras cuentas por pagar",
        "sheets": ["Cuentas por pagar"],
        "category": "pasivos_corrientes"
    },
    "#N18": {
        "title": "Beneficios a los empleados",
        "sheets": ["Provisiones"],  # Mapeada a la pestaña de Provisiones
        "category": "pasivos_no_corrientes"
    },
    "#N19": {
        "title": "Otros pasivos no financieros (Corriente)",
        "sheets": ["Otros pasivos no financieros"],
        "category": "pasivos_corrientes"
    },
    "#N20": {
        "title": "Patrimonio",
        "sheets": ["Patrimonio"],
        "category": "patrimonio"
    },
    "#N21": {
        "title": "Ingresos de actividades ordinarias y costo de ventas",
        "sheets": ["Ingresos Ctos operacion"],
        "category": "resultados"
    },
    "#N22": {
        "title": "Gastos de administración",
        "sheets": ["Gtos Adm"],
        "category": "resultados"
    },
    "#N23": {
        "title": "Diferencia de cambio",
        "sheets": ["DC y Reajustes"],
        "category": "resultados"
    },
    "#N24": {
        "title": "Costos e ingresos financieros",
        "sheets": ["Costos e ingresos Financieros"],
        "category": "resultados"
    },
    "#N25": {
        "title": "Otros ingresos y egresos",
        "sheets": ["Otros gastos por funcion", "Otros ingresos por funcion"],
        "category": "resultados"
    },
    "#N26": {
        "title": "Segmentos de operación",
        "sheets": ["Segmentos"],
        "category": "resultados",
        "consolidated_only": True
    }
}

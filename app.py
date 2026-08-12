import os
import json
from flask import Flask, request, render_template_string

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- MATRIZ DE JERARQUÍA EXTRAÍDA DEL EXCEL DE MACROSEGMENTOS LAMMSA ---
HIERARCHY_DATA = {
  "IIA — Infraestructura Inteligente y Automatización Industrial": {
    "subs": ["IIA-01 — Robótica y Manufactura Avanzada", "IIA-02 — Plataformas IIoT (Plataformas Industriales del Internet de las Cosas)", "IIA-03 — Visión Artificial"],
    "acts": ["333999 — Fabricación de otra maquinaria y equipo para la industria en general", "334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico", "333249 — Fabricación de maquinaria y equipo para otras industrias manufactureras"],
    "sub_map": {
      "IIA-01 — Robótica y Manufactura Avanzada": ["333999 — Fabricación de otra maquinaria y equipo para la industria en general"],
      "IIA-02 — Plataformas IIoT (Plataformas Industriales del Internet de las Cosas)": ["334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico"],
      "IIA-03 — Visión Artificial": ["333249 — Fabricación de maquinaria y equipo para otras industrias manufactureras"]
    }
  },
  "EYU — Energía y Utilidades": {
    "subs": ["EYU-01 — Equipo de Transmisión y Distribución de energía", "EYU-02 — Equipo de Transmisión y Distribución de energía", "EYU-03 — Generación y Distribución de Eneregía", "EYU-04 — Agua", "EYU-05 — Petróleo y Gas"],
    "acts": ["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica", "335311 — Fabricación de motores y generadores eléctricos", "221111 — Generación de electricidad a partir de combustibles fósiles", "221112 — Generación de electricidad a partir de energía hidráulica", "221122 — Distribución de energía eléctrica", "221119 — Generación de electricidad a partir de otro tipo de energía", "221121 — Transmisión de energía eléctrica", "221123 — Comercialización de energía eléctrica", "237133 — Supervisión de construcción de obras de generación y conducción de energía eléctrica y de obras para telecomunicaciones", "237131 — Construcción de obras de generación y conducción de energía eléctrica", "221312 — Captación, tratamiento y suministro de agua (sector público)", "237111 — Construcción de obras para el tratamiento, distribución y suministro de agua y drenaje", "221311 — Captación, tratamiento y suministro de agua (sector privado)", "213111 — Perforación de pozos petroleros y de gas", "211111 — Extracción de petróleo y gas natural asociado", "324110 — Refinación de petróleo", "237122 — Construcción de plantas de refinería y petroquímica", "221210 — Suministro de gas natural por ductos al consumidor final"],
    "sub_map": {
      "EYU-01 — Equipo de Transmisión y Distribución de energía": ["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica"],
      "EYU-02 — Equipo de Transmisión y Distribución de energía": ["335311 — Fabricación de motores y generadores eléctricos"],
      "EYU-03 — Generación y Distribución de Eneregía": ["221111 — Generación de electricidad a partir de combustibles fósiles", "221112 — Generación de electricidad a partir de energía hidráulica", "221122 — Distribución de energía eléctrica", "221119 — Generación de electricidad a partir de otro tipo de energía", "221121 — Transmisión de energía eléctrica", "221123 — Comercialización de energía eléctrica", "237133 — Supervisión de construcción de obras de generación y conducción de energía eléctrica y de obras para telecomunicaciones", "237131 — Construcción de obras de generación y conducción de energía eléctrica"],
      "EYU-04 — Agua": ["221312 — Captación, tratamiento y suministro de agua (sector público)", "237111 — Construcción de obras para el tratamiento, distribución y suministro de agua y drenaje", "221311 — Captación, tratamiento y suministro de agua (sector privado)"],
      "EYU-05 — Petróleo y Gas": ["213111 — Perforación de pozos petroleros y de gas", "211111 — Extracción de petróleo y gas natural asociado", "324110 — Refinación de petróleo", "237122 — Construcción de plantas de refinería y petroquímica", "221210 — Suministro de gas natural por ductos al consumidor final"]
    }
  },
  "ENR — Energías Renovables": {
    "subs": ["ENR-01 — Energía Solar", "ENR-02 — Energía Eólica", "ENR-03 — Almacenamiento de energía y Baterías"],
    "acts": ["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica", "221113 — Generación de electricidad a partir de energía solar", "333610 — Fabricación de motores de combustión interna, turbinas y transmisiones", "221114 — Generación de electricidad a partir de energía eólica", "335910 — Fabricación de acumuladores y pilas"],
    "sub_map": {
      "ENR-01 — Energía Solar": ["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica", "221113 — Generación de electricidad a partir de energía solar"],
      "ENR-02 — Energía Eólica": ["333610 — Fabricación de motores de combustión interna, turbinas y transmisiones", "221114 — Generación de electricidad a partir de energía eólica"],
      "ENR-03 — Almacenamiento de energía y Baterías": ["335910 — Fabricación de acumuladores y pilas"]
    }
  },
  "EEE — Electrónica y Equipos Eléctricos": {
    "subs": ["EEE-01 — Ensamble de PCB (Tarjetas de Circuitos)", "EEE-02 — Semiconductores", "EEE-03 — Equipo eléctrico que incluye Inversores y Convertidores de corriente", "EEE-04 — Conectores", "EEE-05 — Iluminación"],
    "acts": ["334410 — Fabricación de componentes electrónicos", "335999 — Fabricación de otros equipos eléctricos", "335920 — Fabricación de cables de conducción eléctrica", "335120 — Fabricación de lámparas y luminarias"],
    "sub_map": {
      "EEE-01 — Ensamble de PCB (Tarjetas de Circuitos)": ["334410 — Fabricación de componentes electrónicos"],
      "EEE-02 — Semiconductores": ["334410 — Fabricación de componentes electrónicos"],
      "EEE-03 — Equipo eléctrico que incluye Inversores y Convertidores de corriente": ["335999 — Fabricación de otros equipos eléctricos"],
      "EEE-04 — Conectores": ["335920 — Fabricación de cables de conducción eléctrica"],
      "EEE-05 — Iluminación": ["335120 — Fabricación de lámparas y luminarias"]
    }
  },
  "AYD — Aeronáutica y Defensa": {
    "subs": ["AYD-01 — Equipo Original de Aeronaves", "AYD-02 — Sub-ensambles y componentes de Aeroestructuras", "AYD-03 — Sistemas y Motores", "AYD-04 — Interiores de Aeronaves", "AYD-05 — Reparación, Mantenimiento y Overhaul (MRO)", "AYD-06 — Defensa y Seguridad", "AYD-07 — Aviación General", "AYD-08 — Espacial"],
    "acts": ["336410 — Fabricación de equipo aeroespacial", "488190 — Otros servicios relacionados con el transporte aéreo"],
    "sub_map": {
      "AYD-01 — Equipo Original de Aeronaves": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-02 — Sub-ensambles y componentes de Aeroestructuras": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-03 — Sistemas y Motores": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-04 — Interiores de Aeronaves": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-05 — Reparación, Mantenimiento y Overhaul (MRO)": ["488190 — Otros servicios relacionados con el transporte aéreo"],
      "AYD-06 — Defensa y Seguridad": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-07 — Aviación General": ["336410 — Fabricación de equipo aeroespacial"],
      "AYD-08 — Espacial": ["336410 — Fabricación de equipo aeroespacial"]
    }
  },
  "AYT — Automotriz y Transporte": {
    "subs": ["AYT-01 — Vehículos Ligeros (Autos y Camionetas)", "AYT-02 — Vehículos Pesados y Camiones", "AYT-03 — Autobuses y Pasajeros", "AYT-04 — Remolques y Semirremolques", "AYT-05 — Vehículos Especiales (Ambulancias, Limpieza)", "AYT-06 — Ferrocarril y Equipo Ferroviario", "AYT-07 — Motocicletas y Vehículos Ligeros", "AYT-08 — Equipo de Transporte de Carga Especializado", "AYT-09 — Partes y Estampados para Carrocería", "AYT-10 — Motor, Transmisión y Tren Motriz", "AYT-11 — Chasis, Suspensión y Frenos", "AYT-12 — Interiores, Asientos y Acabados", "AYT-13 — Eléctrico, Electrónico e Iluminación Automotriz", "AYT-14 — Llantas, Rines y Hule para Automoción", "AYT-15 — Vidrio y Cristales Automotrices"],
    "acts": ["336110 — Fabricación de automóviles y camiones ligeros", "336210 — Fabricación de carrocerías y remolques", "336310 — Fabricación de motores y sus partes para vehículos automotores", "336320 — Fabricación de equipo eléctrico y electrónico para vehículos automotores", "336330 — Fabricación de partes de sistemas de dirección y de suspensión para vehículos automotores", "336340 — Fabricación de sistemas de frenos para vehículos automotores", "336350 — Fabricación de sistemas de transmisión para vehículos automotores", "336360 — Fabricación de asientos y acesorios interiores para vehículos automotores", "336370 — Fabricación de piezas metálicas troqueladas para vehículos automotores", "336390 — Fabricación de otras partes para vehículos automotores", "326211 — Fabricación de llantas y cámaras", "327216 — Fabricación de vidrio automotriz"],
    "sub_map": {
      "AYT-01 — Vehículos Ligeros (Autos y Camionetas)": ["336110 — Fabricación de automóviles y camiones ligeros"],
      "AYT-02 — Vehículos Pesados y Camiones": ["336110 — Fabricación de automóviles y camiones ligeros"],
      "AYT-04 — Remolques y Semirremolques": ["336210 — Fabricación de carrocerías y remolques"],
      "AYT-09 — Partes y Estampados para Carrocería": ["336370 — Fabricación de piezas metálicas troqueladas para vehículos automotores"],
      "AYT-10 — Motor, Transmisión y Tren Motriz": ["336310 — Fabricación de motores y sus partes para vehículos automotores", "336350 — Fabricación de sistemas de transmisión para vehículos automotores"],
      "AYT-11 — Chasis, Suspensión y Frenos": ["336330 — Fabricación de partes de sistemas de dirección y de suspensión para vehículos automotores", "336340 — Fabricación de sistemas de frenos para vehículos automotores"],
      "AYT-12 — Interiores, Asientos y Acabados": ["336360 — Fabricación de asientos y acesorios interiores para vehículos automotores"],
      "AYT-13 — Eléctrico, Electrónico e Iluminación Automotriz": ["336320 — Fabricación de equipo eléctrico y electrónico para vehículos automotores"],
      "AYT-14 — Llantas, Rines y Hule para Automoción": ["326211 — Fabricación de llantas y cámaras"],
      "AYT-15 — Vidrio y Cristales Automotrices": ["327216 — Fabricación de vidrio automotriz"]
    }
  },
  "MIP — Maquinaria Industrial y Equipo Pesado": {
    "subs": ["MIP-01 — Maquinaria de Construcción y Maquinaria Pesada", "MIP-02 — Maquinaria para Minería y Extracción", "MIP-03 — Maquinaria para Manejo de Materiales y Logística", "MIP-04 — Maquinaria para la Industria del Plástico y Hule", "MIP-05 — Maquinaria para la Industria Textil y del Calzado", "MIP-06 — Maquinaria para la Industria del Papel y Cartón", "MIP-07 — Maquinaria para la Industria de Alimentos y Bebidas", "MIP-08 — Maquinaria Metalmecánica y Máquinas Herramienta", "MIP-09 — Equipos de Elevación y Grúas Industriales"],
    "acts": ["333120 — Fabricación de maquinaria y equipo para la construcción", "333130 — Fabricación de maquinaria y equipo para la minería", "333920 — Fabricación de maquinaria y equipo para levantar y trasladar", "333220 — Fabricación de maquinaria y equipo para la industria del plástico y del hule", "333242 — Fabricación de maquinaria y equipo para la industria editorial y del papel", "333241 — Fabricación de maquinaria y equipo para la industria alimentaria y de bebidas", "333510 — Fabricación de máquinas herramienta para labrar metales"],
    "sub_map": {
      "MIP-01 — Maquinaria de Construcción y Maquinaria Pesada": ["333120 — Fabricación de maquinaria y equipo para la construcción"],
      "MIP-02 — Maquinaria para Minería y Extracción": ["333130 — Fabricación de maquinaria y equipo para la minería"],
      "MIP-03 — Maquinaria para Manejo de Materiales y Logística": ["333920 — Fabricación de maquinaria y equipo para levantar y trasladar"],
      "MIP-04 — Maquinaria para la Industria del Plástico y Hule": ["333220 — Fabricación de maquinaria y equipo para la industria del plástico y del hule"],
      "MIP-06 — Maquinaria para la Industria del Papel y Cartón": ["333242 — Fabricación de maquinaria y equipo para la industria editorial y del papel"],
      "MIP-07 — Maquinaria para la Industria de Alimentos y Bebidas": ["333241 — Fabricación de maquinaria y equipo para la industria alimentaria y de bebidas"],
      "MIP-08 — Maquinaria Metalmecánica y Máquinas Herramienta": ["333510 — Fabricación de máquinas herramienta para labrar metales"]
    }
  },
  "EAR — Equipo Agrícola, Pecuario, Sistemas de Riego y Jardinería": {
    "subs": ["EAR-01 — Maquinaria y Equipo Agrícola", "EAR-02 — Sistemas de Riego y Manejo de Agua Agrícola", "EAR-03 — Equipo Pecuario y Ganadero", "EAR-04 — Equipo de Jardinería y Áreas Verdes"],
    "acts": ["333111 — Fabricación de maquinaria y equipo agrícola", "333112 — Fabricación de cosechadoras y tractores agrícolas"],
    "sub_map": {
      "EAR-01 — Maquinaria y Equipo Agrícola": ["333111 — Fabricación de maquinaria y equipo agrícola", "333112 — Fabricación de cosechadoras y tractores agrícolas"]
    }
  },
  "APG — Agricultura, Pesca, Gandería": {
    "subs": ["APG-01 — Cultivo de Granos y Semillas", "APG-02 — Fruticultura y Hortalizas", "APG-03 — Ganadería y Producción Pecuaria", "APG-04 — Pesca y Acuacultura"],
    "acts": ["111110 — Cultivo de soya, cártamo, girasol y otros granos", "111210 — Cultivo de hortalizas", "112110 — Explotación de bovinos para la producción de carne", "112120 — Explotación de bovinos para la producción de leche"],
    "sub_map": {
      "APG-01 — Cultivo de Granos y Semillas": ["111110 — Cultivo de soya, cártamo, girasol y otros granos"],
      "APG-02 — Fruticultura y Hortalizas": ["111210 — Cultivo de hortalizas"],
      "APG-03 — Ganadería y Producción Pecuaria": ["112110 — Explotación de bovinos para la producción de carne", "112120 — Explotación de bovinos para la producción de leche"]
    }
  },
  "EYC — Edificación y Construcción": {
    "subs": ["EYC-01 — Edificación Residencial y Comercial", "EYC-02 — Infraestructura Vial y Carretera", "EYC-03 — Obras Hidráulicas y Marítimas", "EYC-04 — Estructuras Metálicas para Construcción", "EYC-05 — Instalaciones Especiales en Edificación"],
    "acts": ["236110 — Edificación de vivienda unifamiliar y multifamiliar", "236220 — Edificación no residencial", "237310 — Construcción de vías de comunicación", "238110 — Trabajos de cimentación y estructuras de concreto", "238120 — Montaje de estructuras metálicas"],
    "sub_map": {
      "EYC-01 — Edificación Residencial y Comercial": ["236110 — Edificación de vivienda unifamiliar y multifamiliar", "236220 — Edificación no residencial"],
      "EYC-02 — Infraestructura Vial y Carretera": ["237310 — Construcción de vías de comunicación"],
      "EYC-04 — Estructuras Metálicas para Construcción": ["238120 — Montaje de estructuras metálicas"]
    }
  },
  "FME — Fabriciones Metálicas": {
    "subs": ["FME-01 — Pailería y Soldadura Estructural", "FME-02 — Troquelado, Estampado y Corte de Lámina", "FME-03 — Maquinados de Precisión y Tornería", "FME-04 — Tratamientos Térmicos y Recubrimientos Metálicos", "FME-05 — Forja y Fundición de Metales", "FME-06 — Ensamble Metálico Especializado"],
    "acts": ["332310 — Fabricación de estructuras metálicas y tanques industriales", "332710 — Maquinado de piezas industriales y tornillos", "332810 — Recubrimientos reales y tratamientos térmicos a piezas metálicas", "332110 — Forja y troquelado de piezas metálicas"],
    "sub_map": {
      "FME-01 — Pailería y Soldadura Estructural": ["332310 — Fabricación de estructuras metálicas y tanques industriales"],
      "FME-02 — Troquelado, Estampado y Corte de Lámina": ["332110 — Forja y troquelado de piezas metálicas"],
      "FME-03 — Maquinados de Precisión y Tornería": ["332710 — Maquinado de piezas industriales y tornillos"],
      "FME-04 — Tratamientos Térmicos y Recubrimientos Metálicos": ["332810 — Recubrimientos reales y tratamientos térmicos a piezas metálicas"]
    }
  },
  "PAM — Procesamiento de Alimentos y Manufactura Especializada": {
    "subs": ["PAM-01 — Procesamiento de Carnes y Embutidos", "PAM-02 — Procesamiento de Lácteos y Quesos", "PAM-03 — Panificación y Galletas Industrializadas", "PAM-04 — Procesamiento de Frutas, Vegetales y Bebidas", "PAM-05 — Fabricación de Envases y Empaques Plásticos", "PAM-06 — Manufactura Química y Farmacéutica Especializada"],
    "acts": ["311110 — Elaboración de alimentos para animales", "311210 — Molienda de trigo, maíz y cereales", "311510 — Elaboración de leche y derivados lácteos", "311610 — Matanza, empacado y procesamiento de carne", "311810 — Elaboración de pan y galletas", "312110 — Elaboración de bebidas no alcohólicas y refrescos", "312120 — Elaboración de cerveza y malta", "325110 — Fabricación de petroquímicos básicos", "325410 — Fabricación de productos farmacéuticos", "326110 — Fabricación de bolsas y películas plásticas"],
    "sub_map": {
      "PAM-01 — Procesamiento de Carnes y Embutidos": ["311610 — Matanza, empacado y procesamiento de carne"],
      "PAM-02 — Procesamiento de Lácteos y Quesos": ["311510 — Elaboración de leche y derivados lácteos"],
      "PAM-03 — Panificación y Galletas Industrializadas": ["311810 — Elaboración de pan y galletas"],
      "PAM-04 — Procesamiento de Frutas, Vegetales y Bebidas": ["312110 — Elaboración de bebidas no alcohólicas y refrescos", "312120 — Elaboración de cerveza y malta"],
      "PAM-05 — Fabricación de Envases y Empaques Plásticos": ["326110 — Fabricación de bolsas y películas plásticas"],
      "PAM-06 — Manufactura Química y Farmacéutica Especializada": ["325110 — Fabricación de petroquímicos básicos", "325410 — Fabricación de productos farmacéuticos"]
    }
  },
  "SYM — Salud y Medicina": {
    "subs": ["SYM-01 — Dispositivos Médicos y Equipo Hospitalario", "SYM-02 — Mobiliario Médico y Quirúrgico", "SYM-03 — Instrumental Quirúrgico y de Diagnóstico"],
    "acts": ["339110 — Fabricación de equipo, instrumental y suministros médicos"],
    "sub_map": {"SYM-01 — Dispositivos Médicos y Equipo Hospitalario": ["339110 — Fabricación de equipo, instrumental y suministros médicos"]}
  },
  "MRO — Mantenimiento MRO": {
    "subs": ["MRO-01 — Mantenimiento Industrial General MRO", "MRO-02 — Servicios de Reparación y Overhaul Mecánico", "MRO-03 — Suministros MRO y Consumibles de Planta"],
    "acts": ["541330 — Servicios de ingeniería y servicios relacionados"],
    "sub_map": {"MRO-01 — Mantenimiento Industrial General MRO": ["541330 — Servicios de ingeniería y servicios relacionados"]}
  },
  "ELA — Electrodomésticos, Línea Blanca y Aires Acondicionados Compactos": {
    "subs": ["ELA-01 — Ensamble de Electrodomésticos y Línea Blanca", "ELA-02 — Equipos de Aire Acondicionado y Refrigeración Compacta"],
    "acts": ["335220 — Fabricación de enseres electrodomésticos mayores"],
    "sub_map": {"ELA-01 — Ensamble de Electrodomésticos y Línea Blanca": ["335220 — Fabricación de enseres electrodomésticos mayores"]}
  },
  "MUE — Muebles y Maderas": {
    "subs": ["MUE-01 — Fabricación de Muebles Metálicos y de Madera", "MUE-02 — Muebles para Oficina y Comercio"],
    "acts": ["337120 — Fabricación de muebles para el hogar y oficina"],
    "sub_map": {"MUE-01 — Fabricación de Muebles Metálicos y de Madera": ["337120 — Fabricación de muebles para el hogar y oficina"]}
  },
  "MAR — Marina y Construcción Naval": {
    "subs": ["MAR-01 — Astilleros y Construcción Naval", "MAR-02 — Mantenimiento y Reparación Marítima"],
    "acts": ["336611 — Astilleros y construcción de embarcaciones"],
    "sub_map": {"MAR-01 — Astilleros y Construcción Naval": ["336611 — Astilleros y construcción de embarcaciones"]}
  },
  "CME — Comercio al por menor": {
    "subs": ["CME-01 — Tiendas de Autoservicio y Ferreterías al por Menor"],
    "acts": [], "sub_map": {}
  },
  "CMA — Comercio al por mayor": {
    "subs": ["CMA-01 — Distribuidores Mayoristas e Importadores Industriales"],
    "acts": [], "sub_map": {}
  },
  "MIN — Minería": {
    "subs": ["MIN-01 — Minería de Metales Preciosos e Industriales", "MIN-02 — Extracción de Minerales No Metálicos y Canteras"],
    "acts": ["212210 — Minería de mineral de hierro", "212220 — Minería de oro y plata"],
    "sub_map": {"MIN-01 — Minería de Metales Preciosos e Industriales": ["212210 — Minería de mineral de hierro", "212220 — Minería de oro y plata"]}
  },
  "PYC — Papel, Cartón y productos derivados": {
    "subs": ["PYC-01 — Fabricación de Cajas de Cartón y Empaques", "PYC-02 — Fabricación de Papel y Celulosa"],
    "acts": ["322110 — Fabricación de pulpa, papel y cartón", "322210 — Fabricación de cajas y empaques de cartón corrugado"],
    "sub_map": {
      "PYC-01 — Fabricación de Cajas de Cartón y Empaques": ["322210 — Fabricación de cajas y empaques de cartón corrugado"],
      "PYC-02 — Fabricación de Papel y Celulosa": ["322110 — Fabricación de pulpa, papel y cartón"]
    }
  },
  "OIS — Otras Industrias y Servicios": {
    "subs": ["OIS-01 — Servicios de Ingeniería y Consultoría Técnica", "OIS-02 — Otras Manufacturas Diversas"],
    "acts": ["541330 — Servicios de ingeniería y servicios relacionados"],
    "sub_map": {"OIS-01 — Servicios de Ingeniería y Consultoría Técnica": ["541330 — Servicios de ingeniería y servicios relacionados"]}
  }
}

MACRO_OPTS = list(HIERARCHY_DATA.keys())
REGISTROS_PROSPECTOS = []

# MATRIZ COMPLETA DE TOLLGATES (TG0 A TG13)
TOLLGATES_DATA = {
  "TG0": {
    "objeto": "Lead", "fase": "BD → MO",
    "secciones": [{
      "nombre": "Datos del Contacto",
      "campos": [
        {"id": "nombre", "campo": "Nombre", "tipo": "Texto", "req": True, "ayuda": "Nombre del contacto.", "notas": ""},
        {"id": "apellidos", "campo": "Apellidos", "tipo": "Texto", "req": True, "ayuda": "Apellidos del contacto.", "notas": ""},
        {"id": "empresa", "campo": "Empresa / Razón Social", "tipo": "Texto", "req": True, "ayuda": "Nombre de la empresa del lead.", "notas": ""},
        {"id": "cargo", "campo": "Cargo / Título", "tipo": "Texto", "req": True, "ayuda": "Cargo o posición del contacto.", "notas": ""},
        {"id": "email", "campo": "Email Corporativo", "tipo": "Email", "req": True, "ayuda": "Email corporativo del contacto.", "notas": ""},
        {"id": "telefono", "campo": "Teléfono Contacto", "tipo": "Teléfono", "req": False, "ayuda": "Teléfono directo o móvil.", "notas": ""},
        {"id": "whatsapp", "campo": "WhatsApp (Número)", "tipo": "Teléfono", "req": False, "ayuda": "Número con formato internacional y código de país (ej. +52 1 81 1234 5678).", "notas": ""},
        {"id": "pais_region", "campo": "País / Región", "tipo": "Lista (picklist)", "req": True, "ayuda": "País o región donde opera el lead.", "notas": "", "opts": ["Estados Unidos", "Canadá", "México", "Centroamérica", "España"]}
      ]
    }, {
      "nombre": "Origen y Asignación",
      "campos": [
        {"id": "notas_adicionales", "campo": "Notas Adicionales", "tipo": "Texto largo", "req": False, "ayuda": "Notas generales sobre el lead.", "notas": "NO REQUERIDO"}
      ]
    }]
  },
  "TG1": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Segmentación y Fit",
      "campos": [
        {"id": "macro_segmento", "campo": "Macro Segmento", "tipo": "Lista (picklist)", "req": True, "ayuda": "Macro segmento al que pertenece el lead.", "notas": "", "opts": MACRO_OPTS},
        {"id": "sub_segmento", "campo": "Sub-Segmento", "tipo": "Lista (picklist)", "req": False, "ayuda": "Sub-segmento específico.", "notas": "", "opts": []},
        {"id": "actividad_economica", "campo": "Actividad Económica", "tipo": "Lista (picklist)", "req": True, "ayuda": "Actividad económica detallada SCIAN.", "notas": "", "opts": []},
        {"id": "geografia", "campo": "Geografía", "tipo": "Lista (picklist)", "req": True, "ayuda": "Región geográfica donde opera el cliente.", "notas": "", "opts": ["Norte", "Centro", "Bajio", "Occidente", "Golfo"]},
        {"id": "relevancia", "campo": "Relevancia del Portafolio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Nivel de relevancia del portafolio.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
        {"id": "tamano_empresa", "campo": "Tamaño de Empresa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estimación del tamaño por ventas anuales.", "notas": "", "opts": [
            "Micro — <$10M",
            "Pequeña — $10-50M",
            "Mediana — $50-200M",
            "Grande — $200M-$1B",
            "Enterprise — >$1B"
        ]},
        {"id": "banda", "campo": "Banda Asignada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Banda de clasificación BSV del cliente.", "notas": "si es K o A, entonces es BSV - Normal", "opts": ["K", "A", "B", "C", "D"]}
      ]
    }]
  },
  "TG2": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Engagement y Señales",
      "campos": [
        {"id": "rol_contacto", "campo": "Rol del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rol del contacto en la decisión.", "notas": "", "opts": ["Decision Maker", "Influencer", "Evaluador técnico", "Usuario final", "Desconocido"]}
      ]
    }, {
      "nombre": "Enriquecimiento",
      "campos": [
        {"id": "num_plantas", "campo": "Número de Plantas / Ubicaciones", "tipo": "Número", "req": False, "ayuda": "Número estimado de plantas operativas.", "notas": ""},
        {"id": "contactos_adic", "campo": "Contactos Adicionales Identificados", "tipo": "Texto largo", "req": False, "ayuda": "Nombre, cargo y email de contactos adicionales.", "notas": ""}
      ]
    }]
  },
  "TG3": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Conversión MQL",
      "campos": [
        {"id": "codigo_xz", "campo": "Código XZ", "tipo": "Lista (picklist)", "req": True, "ayuda": "Código de clasificación XZ.", "notas": "Consumo", "opts": ["Directo, Indirecto"]}
      ]
    }, {
      "nombre": "Notas de Transferencia",
      "campos": [
        {"id": "puntos_interes", "campo": "Puntos de Interés Detectados", "tipo": "Texto largo", "req": True, "ayuda": "Hipótesis de necesidades o pain points.", "notas": ""},
        {"id": "calidad_contacto", "campo": "Calidad del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación subjetiva de la calidad del contacto.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
        {"id": "comentarios_ventas", "campo": "Comentarios Adicionales para Ventas", "tipo": "Texto largo", "req": False, "ayuda": "Notas adicionales para Ventas.", "notas": ""}
      ]
    }]
  },
  "TG4": {
    "objeto": "Lead", "fase": "MO → WO",
    "secciones": [{
      "nombre": "Validación ERP",
      "campos": [
        {"id": "ventas_12m", "campo": "Ventas Últimos 12 Meses", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rango de ventas de los últimos 12 meses.", "notas": "", "opts": ["Sin historial", "$0", "<$20k"]},
        {"id": "actividad_lammsa", "campo": "Actividad Lammsa Previa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si tiene historial con productos Sommer.", "notas": "", "opts": ["Nunca ha comprado", "Histórico inactivo", "Activo — producto Sommer"]}
      ]
    }, {
      "nombre": "Clasificación Comercial",
      "campos": [
        {"id": "clasif_comercial", "campo": "Clasificación Comercial", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial del lead.", "notas": "", "opts": ["K, A, B, C, D"]},
        {"id": "justif_clasif", "campo": "Justificación de Clasificación", "tipo": "Texto largo", "req": True, "ayuda": "Evidencia de la lógica aplicada.", "notas": ""}
      ]
    }, {
      "nombre": "Asignación y SLA",
      "campos": [
        {"id": "osp_asignado", "campo": "OSP Asignado", "tipo": "Texto", "req": True, "ayuda": "Nombre del OSP que recibe la cuenta.", "notas": ""},
        {"id": "gerente_area", "campo": "Gerente de Área", "tipo": "Texto", "req": True, "ayuda": "Nombre del Gerente responsable.", "notas": ""},
        {"id": "fecha_asignacion", "campo": "Fecha de Asignación", "tipo": "Fecha", "req": True, "ayuda": "Fecha de asignación formal.", "notas": ""},
        {"id": "fecha_comp_1ra_inter", "campo": "Fecha Compromiso 1ra Interacción", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para primera interacción.", "notas": "Convierte a cuenta"}
      ]
    }]
  },
  "TG5": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos de la WO",
      "campos": [
        {"id": "tipo_oportunidad", "campo": "Tipo de Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial.", "notas": "", "opts": ["XR — Reactivación", "XP — Prospecto", "XS — Cross Sell Sommer"]},
        {"id": "aoi_estimado", "campo": "AOI Estimado (%)", "tipo": "Porcentaje (%)", "req": True, "ayuda": "Porcentaje de ahorro o ingreso estimado.", "notas": ""},
        {"id": "primera_inter_ok", "campo": "Primera Interacción Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se completó la primera interacción.", "notas": "Enviar Template de Agradecimiento", "opts": ["Sí", "No — pendiente"]},
        {"id": "fecha_1ra_inter", "campo": "Fecha de Primera Interacción", "tipo": "Fecha", "req": False, "ayuda": "Fecha real de la primera interacción.", "notas": ""},
        {"id": "formato_1ra_inter", "campo": "Formato de Primera Interacción", "tipo": "Lista (picklist)", "req": False, "ayuda": "Canal o formato utilizado.", "notas": "", "opts": ["Presencial — planta", "Presencial — oficina", "Virtual"]}
      ]
    }, {
      "nombre": "Plan BSV",
      "campos": [
        {"id": "estado_plan_bsv", "campo": "Estado del Plan BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado del Plan BSV para la oportunidad.", "notas": "", "opts": ["Sí — completado", "En proceso", "No — pendiente"]}
      ]
    }]
  },
  "TG6": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Plan BSV",
      "campos": [
        {"id": "fecha_est_plan_bsv", "campo": "Fecha Estimada Presentación Plan BSV", "tipo": "Fecha", "req": False, "ayuda": "Fecha estimada para presentar el Plan BSV.", "notas": ""},
        {"id": "notas_plan_bsv", "campo": "Notas del Plan BSV", "tipo": "Texto largo", "req": False, "ayuda": "Consideraciones iniciales del Plan BSV.", "notas": ""}
      ]
    }, {
      "nombre": "Decisión GO / NO-GO",
      "campos": [
        {"id": "decision_go", "campo": "Decisión GO / NO-GO", "tipo": "Lista (picklist)", "req": True, "ayuda": "Decisión formal de continuidad.", "notas": "Flujo de Aprobación Gerencia", "opts": ["GO — continuar con investigación BSV", "NO-GO — salida", "TICKLER — re-evaluar en fecha futura"]},
        {"id": "justif_decision", "campo": "Justificación de la Decisión", "tipo": "Texto largo", "req": True, "ayuda": "Justificación obligatoria.", "notas": ""},
        {"id": "siguiente_paso", "campo": "Siguiente Paso Concreto", "tipo": "Texto", "req": True, "ayuda": "Siguiente acción específica.", "notas": ""},
        {"id": "fecha_sig_paso", "campo": "Fecha del Siguiente Paso", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para el siguiente paso.", "notas": ""}
      ]
    }, {
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_i_score", "campo": "MEDDICC I — Score (Dolor Implicado)", "tipo": "Número", "req": True, "ayuda": "Score Implicated Pain (1-10). Gate: >=4.", "notas": ""},
        {"id": "meddicc_i_notas", "campo": "MEDDICC I — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Descripción del dolor y costo de inacción.", "notas": ""},
        {"id": "meddicc_m_score", "campo": "MEDDICC M — Score (Métricas)", "tipo": "Número", "req": True, "ayuda": "Score Metrics (1-10). Gate: >=5.", "notas": ""},
        {"id": "meddicc_m_notas", "campo": "MEDDICC M — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Métricas cuantificadas de impacto.", "notas": ""},
        {"id": "meddicc_c2_score", "campo": "MEDDICC C2 — Score (Competencia)", "tipo": "Número", "req": True, "ayuda": "Score Competition (1-10).", "notas": ""},
        {"id": "meddicc_c2_notas", "campo": "MEDDICC C2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Proveedor actual y ventajas de LAMMSA.", "notas": ""}
      ]
    }]
  },
  "TG7": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_c1_score", "campo": "MEDDICC C1 — Score (Campeón)", "tipo": "Número", "req": True, "ayuda": "Score Champion (1-10).", "notas": "Agregar en contacto"},
        {"id": "meddicc_c1_notas", "campo": "MEDDICC C1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre del Campeón y evidencia.", "notas": "Requiere interacción previa"},
        {"id": "meddicc_d1_score", "campo": "MEDDICC D1 — Score (Criterios de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Criteria (1-10).", "notas": ""},
        {"id": "meddicc_d1_notas", "campo": "MEDDICC D1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Criterios de evaluación del cliente.", "notas": ""}
      ]
    }]
  },
  "TG8": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_e_score", "campo": "MEDDICC E — Score (Comprador Económico)", "tipo": "Número", "req": True, "ayuda": "Score Economic Buyer (1-10).", "notas": "Agregar en contacto"},
        {"id": "meddicc_e_notas", "campo": "MEDDICC E — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y nivel de soporte del Economic Buyer.", "notas": "Requiere interacción previa"},
        {"id": "meddicc_d2_score", "campo": "MEDDICC D2 — Score (Proceso de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Process (1-10).", "notas": ""},
        {"id": "meddicc_d2_notas", "campo": "MEDDICC D2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos del proceso de decisión y timeline.", "notas": ""},
        {"id": "meddicc_p_score", "campo": "MEDDICC P — Score (Proceso de Contratación)", "tipo": "Número", "req": True, "ayuda": "Score Paper Process (1-10).", "notas": ""},
        {"id": "meddicc_p_notas", "campo": "MEDDICC P — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos de revisión legal y firmas.", "notas": ""}
      ]
    }]
  },
  "TG9": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Investigación BSV",
      "campos": [
        {"id": "resumen_bsv", "campo": "Resumen Ejecutivo BSV", "tipo": "Texto largo", "req": True, "ayuda": "Resumen de la situación y propuesta.", "notas": ""},
        {"id": "mapa_stakeholders", "campo": "Mapa de Stakeholders y Coaches", "tipo": "Texto largo", "req": True, "ayuda": "Mapa de actores clave.", "notas": ""},
        {"id": "hipotesis_fit", "campo": "Hipótesis de Fit de Negocio", "tipo": "Texto largo", "req": True, "ayuda": "Cómo LAMMSA resuelve el problema.", "notas": ""},
        {"id": "aoi_cuantificado", "campo": "AOI Cuantificado ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "Valor en dólares del AOI.", "notas": ""},
        {"id": "analisis_incumbentes", "campo": "Análisis de Incumbentes", "tipo": "Texto largo", "req": True, "ayuda": "Análisis de competidores actuales.", "notas": ""},
        {"id": "posicionamiento_diferenciador", "campo": "Posicionamiento Diferenciador LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Argumento diferenciador de LAMMSA.", "notas": ""}
      ]
    }, {
      "nombre": "Aprobación Gerencial",
      "campos": [
        {"id": "aprobacion_summit", "campo": "Aprobación para Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación para el Summit.", "notas": "Subir deck BSV", "opts": ["Aprobado para Summit", "Pendiente de revisión", "No aprobado"]},
        {"id": "fecha_rev_gerencial", "campo": "Fecha de Revisión Gerencial", "tipo": "Fecha", "req": False, "ayuda": "Fecha de revisión gerencial.", "notas": ""},
        {"id": "notas_aprobacion", "campo": "Notas de Aprobación", "tipo": "Texto largo", "req": False, "ayuda": "Notas del Gerente de Área.", "notas": ""}
      ]
    }]
  },
  "TG10": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos del Summit",
      "campos": [
        {"id": "fecha_bsv_summit", "campo": "Fecha del BSV Summit", "tipo": "Fecha", "req": True, "ayuda": "Fecha acordada para el Summit.", "notas": "Mail de confirmación requerido"},
        {"id": "participantes_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y cargo de asistentes del cliente.", "notas": ""},
        {"id": "participantes_lammsa", "campo": "Participantes LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Representantes de LAMMSA presentes.", "notas": ""}
      ]
    }]
  },
  "TG11": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos del Summit",
      "campos": [
        {"id": "deck_presentado", "campo": "Deck Presentado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el deck fue presentado.", "notas": "", "opts": ["Sí", "No — pendiente"]},
        {"id": "roundtable_ejecutado", "campo": "Round-Table Ejecutado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se ejecutó el round-table.", "notas": "", "opts": ["Sí", "Parcial", "No"]},
        {"id": "acuerdos_summit", "campo": "Acuerdos y Próximos Pasos del Summit", "tipo": "Texto largo", "req": True, "ayuda": "Acuerdos formalizados.", "notas": "Mail de agradecimiento requerido"}
      ]
    }, {
      "nombre": "Resultado del Summit",
      "campos": [
        {"id": "resultado_summit", "campo": "Resultado del Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado formal del Summit.", "notas": "", "opts": ["Compromiso verbal — proceden", "Compromiso escrito — NDA / LOI firmado", "Interés sin compromiso — más información", "Sin compromiso — no avanzan"]},
        {"id": "notas_resultado_summit", "campo": "Notas del Resultado del Summit", "tipo": "Texto largo", "req": False, "ayuda": "Notas y objeciones del Summit.", "notas": ""}
      ]
    }, {
      "nombre": "Alcance de la Solución",
      "campos": [
        {"id": "plantas_scope", "campo": "Plantas en Scope", "tipo": "Texto", "req": True, "ayuda": "Plantas incluidas en el alcance.", "notas": ""},
        {"id": "categorias_solucion", "campo": "Categorías Incluidas en la Solución", "tipo": "Texto largo", "req": True, "ayuda": "Categorías de producto/servicio.", "notas": ""},
        {"id": "modelo_servicio", "campo": "Modelo de Servicio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Modelo de servicio propuesto.", "notas": "", "opts": ["Full service — inventario LAMMSA", "Consignación", "Mixto", "A definir"]},
        {"id": "plan_implementacion", "campo": "Plan de Implementación por Fases", "tipo": "Texto largo", "req": True, "ayuda": "Entregables, fechas y hitos.", "notas": ""}
      ]
    }, {
      "nombre": "Economía de la Propuesta",
      "campos": [
        {"id": "calculadora_aoi_ok", "campo": "Calculadora AOI Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Validación por Finanzas.", "notas": "", "opts": ["Sí — validada", "No — pendiente"]},
        {"id": "aoi_proyectado_ano1", "campo": "AOI Proyectado Año 1 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado primer año ($USD).", "notas": ""},
        {"id": "aoi_proyectado_anos2_3", "campo": "AOI Proyectado Años 2-3 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado años 2-3 ($USD).", "notas": ""},
        {"id": "margen_vs_floor", "campo": "Margen vs Floor Mínimo", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación de margen vs floor.", "notas": "Bajo el floor no procede", "opts": ["Sobre el floor — aprobado", "En el floor — requiere excepción", "Bajo el floor — no procede"]}
      ]
    }]
  },
  "TG12": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Aprobaciones Internas",
      "campos": [
        {"id": "aprobacion_ops", "campo": "Aprobación Ops / Supply Chain", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Operaciones.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
        {"id": "aprobacion_finanzas", "campo": "Aprobación Finanzas", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Finanzas.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
        {"id": "propuesta_aprobada_int", "campo": "Propuesta Internamente Aprobada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indicador global de aprobaciones.", "notas": "", "opts": ["Sí — aprobación completa", "No — pendiente aprobaciones"]},
        {"id": "fecha_presentacion_cliente", "campo": "Fecha de Presentación al Cliente", "tipo": "Fecha", "req": True, "ayuda": "Fecha confirmada de presentación.", "notas": "Mail de confirmación requerido"}
      ]
    }]
  },
  "TG13": {
    "objeto": "Opportunity", "fase": "WO → SO",
    "secciones": [{
      "nombre": "Presentación y Negociación",
      "campos": [
        {"id": "fecha_presentacion_prop", "campo": "Fecha de Presentación de Propuesta", "tipo": "Fecha", "req": True, "ayuda": "Fecha real de presentación.", "notas": "Mail de agradecimiento requerido"},
        {"id": "participantes_pres_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Asistentes a la presentación.", "notas": ""},
        {"id": "ce_presente", "campo": "¿Economic Buyer Presente?", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el CE estuvo presente.", "notas": "", "opts": ["Sí", "No — representado por Campeón", "No — sin representación ejecutiva"]},
        {"id": "objeciones_manejo", "campo": "Objeciones Planteadas y Manejo", "tipo": "Texto largo", "req": False, "ayuda": "Objeciones y respuestas.", "notas": ""},
        {"id": "estado_negociacion", "campo": "Estado de la Negociación", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado de negociación.", "notas": "", "opts": ["Sin objeciones — procede", "Objeciones menores en proceso", "Objeciones mayores — negociación activa", "Stall — cliente no responde"]}
      ]
    }, {
      "nombre": "Forecast",
      "campos": [
        {"id": "forecast_categoria", "campo": "Categoría de Forecast BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Categoría de forecast BSV.", "notas": "Commit requiere CE + Campeón + Paper Process", "opts": ["Pipeline", "Best Case", "Commit Candidato", "Commit", "Ganado — PO emitida"]}
      ]
    }, {
      "nombre": "Resultado Final",
      "campos": [
        {"id": "resultado_final", "campo": "Resultado Final de la Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado final.", "notas": "", "opts": ["GANADO — PO emitida", "GANADO — Contrato firmado", "TICKLER — Postergado con fecha", "PERDIDO — XO definitivo"]},
        {"id": "notas_resultado_final", "campo": "Notas del Resultado Final", "tipo": "Texto largo", "req": False, "ayuda": "Notas y aprendizajes del cierre.", "notas": ""}
      ]
    }]
  }
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Salesforce — Prospectos BSV</title>
    <style>
        :root {
            --sf-brand: #0176d3;
            --sf-brand-dark: #005fb2;
            --sf-path-blue: #00396b;
            --sf-green: #2e844a;
            --sf-bg: #b0c4df;
            --sf-border: #dddbda;
            --sf-text-main: #181818;
            --sf-text-muted: #514f4d;
            --sf-required: #ea001e;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: var(--sf-bg);
            margin: 0; padding: 0;
            color: var(--sf-text-main);
        }

        .sf-global-header {
            background-color: #ffffff; border-bottom: 1px solid var(--sf-border);
            padding: 0 16px; display: flex; align-items: center; justify-content: space-between; height: 50px;
        }
        .sf-app-launcher { display: flex; align-items: center; gap: 12px; }
        .sf-waffle { display: grid; grid-template-columns: repeat(3, 4px); gap: 3px; cursor: pointer; }
        .sf-waffle div { width: 4px; height: 4px; background-color: var(--sf-brand); border-radius: 1px; }
        .sf-app-title { font-weight: 700; font-size: 16px; color: #181818; }
        
        .sf-nav-menu { display: flex; align-items: center; gap: 4px; height: 100%; }
        .sf-nav-item {
            padding: 0 12px; height: 100%; display: flex; align-items: center;
            font-size: 13px; color: #514f4d; cursor: pointer; text-decoration: none; border-bottom: 3px solid transparent;
        }
        .sf-nav-item.active { color: var(--sf-brand); font-weight: 700; border-bottom-color: var(--sf-brand); }

        .sf-container {
            max-width: 1400px; margin: 12px auto; background: #ffffff;
            border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            overflow: hidden; border: 1px solid var(--sf-border);
        }

        /* VISTA 1: TABLA VISTOS RECIENTEMENTE */
        .sf-list-header {
            padding: 16px 24px; background: #ffffff; border-bottom: 1px solid var(--sf-border);
            display: flex; align-items: center; justify-content: space-between;
        }
        .sf-list-title { display: flex; align-items: center; gap: 12px; }
        .sf-lead-icon {
            width: 32px; height: 32px; background-color: #4bca81; border-radius: 4px;
            display: flex; align-items: center; justify-content: center; color: white; font-size: 18px;
        }
        .sf-btn-nuevo {
            background-color: var(--sf-brand); color: #ffffff; border: 1px solid var(--sf-brand);
            padding: 8px 20px; border-radius: 4px; font-weight: 700; font-size: 13px; cursor: pointer;
        }
        .sf-btn-nuevo:hover { background-color: var(--sf-brand-dark); }
        .sf-btn-sub {
            background-color: #ffffff; color: var(--sf-brand); border: 1px solid var(--sf-border);
            padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer;
        }
        .sf-btn-action-icon {
            background: none; border: none; font-size: 15px; cursor: pointer; padding: 4px 6px; border-radius: 4px; transition: background 0.2s;
        }
        .sf-btn-action-icon:hover { background: #e0e0e0; }

        .sf-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .sf-table th { background-color: #fafafa; border-bottom: 2px solid var(--sf-border); padding: 10px 14px; text-align: left; color: #514f4d; font-weight: 700; }
        .sf-table td { padding: 12px 14px; border-bottom: 1px solid var(--sf-border); color: #181818; }
        .sf-table tr:hover { background-color: #f8f9fb; }

        /* VISTA 2: DETALLE / CAPTURA SPLIT */
        .sf-detail-top-bar {
            padding: 12px 24px; background: #ffffff; border-bottom: 1px solid var(--sf-border);
            display: flex; align-items: center; justify-content: space-between;
        }
        .sf-highlights-grid {
            display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px;
            padding: 12px 24px; background-color: #fafafa; border-bottom: 1px solid var(--sf-border);
        }
        .sf-highlight-item span { display: block; font-size: 11px; color: var(--sf-text-muted); font-weight: 600; }
        .sf-highlight-item strong { font-size: 13px; color: var(--sf-text-main); }

        .sf-path-bar { display: flex; background-color: #f3f3f3; padding: 8px 16px; border-bottom: 1px solid var(--sf-border); overflow-x: auto; white-space: nowrap; }
        
        .sf-chevron {
            padding: 8px 16px; font-size: 12px; font-weight: 700; text-align: center; border: 1px solid var(--sf-border); border-radius: 4px; margin-right: 6px;
            transition: all 0.2s ease;
        }
        .sf-chevron.active {
            background-color: var(--sf-brand); color: #ffffff; border-color: var(--sf-brand); cursor: pointer;
        }
        .sf-chevron.completed {
            background-color: var(--sf-green); color: #ffffff; border-color: var(--sf-green); cursor: pointer;
        }
        .sf-chevron.disabled {
            background-color: #e0e0e0; color: #888888; border-color: #cccccc; cursor: not-allowed; opacity: 0.45; pointer-events: none;
        }

        .sf-split-layout { display: grid; grid-template-columns: 7fr 3fr; background-color: #b0c4df; gap: 12px; padding: 12px; }
        @media (max-width: 992px) { .sf-split-layout { grid-template-columns: 1fr; } }

        .sf-main-col { background: #ffffff; border-radius: 4px; border: 1px solid var(--sf-border); }
        .sf-side-col { display: flex; flex-direction: column; gap: 12px; }

        .sf-tabs { display: flex; border-bottom: 1px solid var(--sf-border); background: #ffffff; padding-left: 16px; }
        .sf-tab { padding: 12px 20px; font-size: 13px; font-weight: 600; color: #514f4d; cursor: pointer; border-bottom: 3px solid transparent; }
        .sf-tab.active { color: var(--sf-brand); border-bottom-color: var(--sf-brand); }

        .sf-card-section { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; margin: 16px; overflow: hidden; }
        .sf-card-header { background: #f3f3f3; padding: 10px 16px; font-size: 13px; font-weight: 700; border-bottom: 1px solid var(--sf-border); }
        .sf-field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 24px; padding: 16px; }

        .sf-field-group { display: flex; flex-direction: column; }
        .sf-label { font-size: 12px; font-weight: 600; color: #514f4d; margin-bottom: 4px; }
        .sf-req { color: var(--sf-required); font-weight: bold; }
        .sf-input, .sf-select, .sf-textarea { padding: 7px 10px; border: 1px solid var(--sf-border); border-radius: 4px; font-size: 13px; box-sizing: border-box; width: 100%; }
        .sf-input:disabled, .sf-select:disabled, .sf-textarea:disabled { background-color: #f3f3f3; color: #555555; cursor: not-allowed; border-color: #dddbda; }
        .sf-help-text { font-size: 11px; color: #514f4d; margin-top: 3px; }

        .sf-side-card { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; padding: 14px; }
        .sf-drop-box { border: 2px dashed var(--sf-border); border-radius: 4px; padding: 20px; text-align: center; background: #fafafa; margin-top: 8px; }

        .alert-success { background-color: #d4edda; color: #155724; padding: 12px 16px; border-radius: 4px; margin: 12px; border: 1px solid #c3e6cb; font-size: 13px; font-weight:600; }
        
        .badge-read-only { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; }
        .badge-edit-mode { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; }
        
        .highlight-yellow { background-color: #fff9c4 !important; }
    </style>
</head>
<body>

<div class="sf-global-header">
    <div class="sf-app-launcher">
        <div class="sf-waffle"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>
        <span class="sf-app-title">Ventas</span>
    </div>
    <div class="sf-nav-menu">
        <a class="sf-nav-item">Inicio</a>
        <a class="sf-nav-item">Cuentas ▾</a>
        <a class="sf-nav-item">Contactos ▾</a>
        <a class="sf-nav-item active">Prospectos ▾</a>
        <a class="sf-nav-item">Oportunidades ▾</a>
        <a class="sf-nav-item">Tareas ▾</a>
        <a class="sf-nav-item">Calendario</a>
        <a class="sf-nav-item">Reportes ▾</a>
    </div>
    <div style="font-size:12px; color:#514f4d;">Salesforce CRM BSV</div>
</div>

<div class="sf-container">

    <!-- VISTA 1: TABLA VISTOS RECIENTEMENTE (INICIO) -->
    <div id="vista-lista" style="display: {% if mostrar_detalle %}none{% else %}block{% endif %};">
        <div class="sf-list-header">
            <div class="sf-list-title">
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;">Prospectos</span>
                    <h2 style="margin:0; font-size:18px;">Vistos recientemente ▾</h2>
                </div>
            </div>
            <div>
                <form method="POST" style="display:inline;">
                    <input type="hidden" name="action_type" value="nuevo">
                    <button type="submit" class="sf-btn-nuevo">+ Nuevo</button>
                </form>
            </div>
        </div>

        {% if mensaje and not mostrar_detalle %}
            <div class="alert-success">{{ mensaje }}</div>
        {% endif %}

        <div style="overflow-x: auto;">
            <table class="sf-table">
                <thead>
                    <tr>
                        <th width="30"><input type="checkbox"></th>
                        <th>Nombre completo (Ver Lectura)</th>
                        <th>Cargo</th>
                        <th>Compañía</th>
                        <th>Teléfono</th>
                        <th>Email</th>
                        <th>Estado</th>
                        <th style="text-align: center; width: 110px;">Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    {% if registros %}
                        {% for r in registros %}
                            {% set datos = r.datos %}
                            {% set nombre_comp = (datos.get('nombre', '') + ' ' + datos.get('apellidos', '')).strip() %}
                            <tr>
                                <td><input type="checkbox"></td>
                                
                                <td>
                                    <form method="POST" style="display:inline;">
                                        <input type="hidden" name="action_type" value="ver_lectura">
                                        <input type="hidden" name="prospecto_id" value="{{ loop.index0 }}">
                                        <button type="submit" style="background:none; border:none; color:var(--sf-brand); font-weight:600; cursor:pointer; font-size:13px; text-decoration:underline; padding:0;">
                                            {{ nombre_comp or '— Sin registrar —' }}
                                        </button>
                                    </form>
                                </td>
                                <td>{{ datos.get('cargo', '—') }}</td>
                                <td>{{ datos.get('empresa', '—') }}</td>
                                <td>{{ datos.get('telefono', '—') }}</td>
                                <td>{{ datos.get('email', '—') }}</td>
                                <td><span style="background:#eef4fe; color:#0176d3; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600;">Nuevo</span></td>
                                
                                <td style="text-align: center; white-space: nowrap;">
                                    <form method="POST" style="display:inline;">
                                        <input type="hidden" name="action_type" value="editar">
                                        <input type="hidden" name="prospecto_id" value="{{ loop.index0 }}">
                                        <button type="submit" class="sf-btn-action-icon" title="Editar prospecto">🖊️</button>
                                    </form>
                                    
                                    <form method="POST" style="display:inline;" onsubmit="return confirm('¿Está seguro de borrar el prospecto {{ nombre_comp }}?');">
                                        <input type="hidden" name="action_type" value="eliminar">
                                        <input type="hidden" name="prospecto_id" value="{{ loop.index0 }}">
                                        <button type="submit" class="sf-btn-action-icon" title="Borrar prospecto">🗑️</button>
                                    </form>
                                </td>
                            </tr>
                        {% endfor %}
                    {% else %}
                        <tr>
                            <td colspan="8" style="text-align: center; padding: 35px; color: #514f4d;">
                                No hay prospectos registrados aún. Haz clic en el botón <strong>"+ Nuevo"</strong> para agregar un registro.
                            </td>
                        </tr>
                    {% endif %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- VISTA 2: DETALLE / CAPTURA SPLIT (70% / 30%) -->
    <div id="vista-detalle" style="display: {% if mostrar_detalle %}block{% else %}none{% endif %};">
        
        <div class="sf-detail-top-bar">
            <div style="display:flex; align-items:center; gap:12px;">
                <form method="POST" style="display:inline;">
                    <input type="hidden" name="action_type" value="volver_lista">
                    <button type="submit" class="sf-btn-sub">← Volver a la Lista</button>
                </form>
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;" id="header-objeto-fase">Objeto SF: Lead | Fase: BD → MO</span>
                    <h1 id="dyn-lead-title" style="margin:0; font-size:18px;">
                        {% set d_curr = registro_actual.datos if registro_actual else {} %}
                        {% set n_curr = (d_curr.get('nombre', '') + ' ' + d_curr.get('apellidos', '')).strip() %}
                        {{ n_curr or '— Sin registrar —' }}
                    </h1>
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                {% if modo_lectura %}
                    <span class="badge-read-only">🔒 Modo Lectura (Solo Consulta)</span>
                    <form method="POST" style="display:inline;">
                        <input type="hidden" name="action_type" value="cambiar_a_edicion">
                        <input type="hidden" name="prospecto_id" value="{{ prospecto_id }}">
                        <button type="submit" class="sf-btn-sub">✏️ Pasar a Modo Edición</button>
                    </form>
                {% else %}
                    <span class="badge-edit-mode">✏️ Modo Edición</span>
                {% endif %}
            </div>
        </div>

        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Empresa / Razón Social</span>
                <strong id="dyn-empresa">{{ d_curr.get('empresa', '—') }}</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Cargo</span>
                <strong id="dyn-cargo">{{ d_curr.get('cargo', '—') }}</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Teléfono</span>
                <strong id="dyn-telefono">{{ d_curr.get('telefono', '—') }}</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Email</span>
                <strong id="dyn-email">{{ d_curr.get('email', '—') }}</strong>
            </div>
        </div>

        <!-- BARRA DE CHEVRONS CON NAVEGACIÓN -->
        <div class="sf-path-bar">
            {% for tg_key in tg_keys %}
                {% set tg_idx = loop.index0 %}
                <div class="sf-chevron {% if tg_key == active_tg %}active{% elif modo_lectura or tg_idx < unlocked_idx %}completed{% else %}disabled{% endif %}"
                     id="tab-btn-{{ tg_key }}"
                     onclick="{% if modo_lectura or tg_idx <= unlocked_idx %}activarTollgate('{{ tg_key }}', {{ tg_idx }}){% else %}return false;{% endif %}">
                    {% if modo_lectura or tg_idx < unlocked_idx %}✓ {% endif %}{{ tg_key }}
                </div>
            {% endfor %}
        </div>

        <!-- FORMULARIO PRINCIPAL -->
        <form method="POST" id="form-prospecto" onsubmit="prepararFormularioParaEnvio()">
            <input type="hidden" name="action_type" value="avanzar">
            <input type="hidden" name="prospecto_id" value="{{ prospecto_id }}">
            <input type="hidden" name="current_active_tg" id="current_active_tg" value="{{ active_tg }}">
            <input type="hidden" name="unlocked_idx" id="unlocked_idx" value="{{ unlocked_idx }}">

            <div class="sf-split-layout">
                
                <div class="sf-main-col">
                    <div class="sf-tabs">
                        <div class="sf-tab active">Detalles de Captura</div>
                        <div class="sf-tab">Actividad</div>
                        <div class="sf-tab">Chatter</div>
                    </div>

                    {% if mensaje and mostrar_detalle %}
                        <div class="alert-success">{{ mensaje }}</div>
                    {% endif %}

                    {% for tg_key, tg_val in tg_data.items() %}
                        <div id="pantalla-{{ tg_key }}" class="tg-screen" style="{% if tg_key != active_tg %}display:none;{% endif %}">
                            {% for seccion in tg_val['secciones'] %}
                                <div class="sf-card-section">
                                    <div class="sf-card-header">▼ {{ seccion['nombre'] }}</div>
                                    <div class="sf-field-grid">
                                        {% for field in seccion['campos'] %}
                                            {% set f_val = d_curr.get(field['id'], '') %}
                                            <div class="sf-field-group" {% if field['tipo'] == 'Texto largo' %}style="grid-column: span 2;"{% endif %}>
                                                <div class="sf-label">
                                                    {% if field['req'] and not modo_lectura %}<span class="sf-req">*</span>{% endif %}{{ field['campo'] }}
                                                </div>

                                                {% if field['tipo'] == 'Lista (picklist)' %}
                                                    <select name="{{ field['id'] }}" 
                                                            class="sf-select {% if field['id'] == 'actividad_economica' and not modo_lectura %}highlight-yellow{% endif %}"
                                                            data-saved-val="{{ f_val }}"
                                                            data-req="{% if field['req'] %}true{% else %}false{% endif %}"
                                                            {% if field['id'] == 'macro_segmento' %}onchange="actualizarCascadaTG1(true)"{% endif %}
                                                            {% if field['id'] == 'sub_segmento' %}onchange="actualizarCascadaTG1(false)"{% endif %}
                                                            {% if modo_lectura %}disabled{% endif %}>
                                                        <option value="">--Seleccione {{ field['campo'] }}--</option>
                                                        {% if field['opts'] %}
                                                            {% for opt in field['opts'] %}
                                                                <option value="{{ opt }}" {% if f_val == opt %}selected{% endif %}>{{ opt }}</option>
                                                            {% endfor %}
                                                        {% endif %}
                                                    </select>
                                                {% elif field['tipo'] == 'Texto largo' %}
                                                    <textarea name="{{ field['id'] }}" class="sf-textarea" rows="3" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %}>{{ f_val }}</textarea>
                                                {% elif field['tipo'] == 'Fecha' %}
                                                    <input type="date" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %}>
                                                {% elif field['tipo'] == 'Email' %}
                                                    <input type="email" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] == 'Teléfono' %}
                                                    <input type="tel" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" placeholder="{% if field['id'] == 'whatsapp' %}+52 1 81 1234 5678{% else %}+52 81 0000 0000{% endif %}" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                                    <input type="number" step="any" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %}>
                                                {% else %}
                                                    <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% endif %}

                                                {% if field['ayuda'] %}
                                                    <span class="sf-help-text">{{ field['ayuda'] }}</span>
                                                {% endif %}
                                            </div>
                                        {% endfor %}
                                    </div>
                                </div>
                            {% endfor %}
                        </div>
                    {% endfor %}

                    <!-- ACCIONES INFERIORES -->
                    <div style="padding: 16px; text-align: right; background: #ffffff; border-top: 1px solid var(--sf-border);">
                        {% if modo_lectura %}
                            <form method="POST" style="display:inline;">
                                <input type="hidden" name="action_type" value="volver_lista">
                                <button type="submit" class="sf-btn-sub" style="margin-right:8px;">Cerrar / Volver a la Lista</button>
                            </form>
                            <form method="POST" style="display:inline;">
                                <input type="hidden" name="action_type" value="cambiar_a_edicion">
                                <input type="hidden" name="prospecto_id" value="{{ prospecto_id }}">
                                <button type="submit" class="sf-btn-nuevo">✏️ Pasar a Modo Edición</button>
                            </form>
                        {% else %}
                            <button type="button" class="sf-btn-sub" onclick="volverALista()" style="margin-right: 8px;">Cancelar</button>
                            <button type="submit" class="sf-btn-nuevo">
                                {% if active_tg == 'TG13' %}Avanzar y Finalizar{% else %}Avanzar ➔{% endif %}
                            </button>
                        {% endif %}
                    </div>
                </div>

                <!-- BARRA LATERAL DERECHA (30%) -->
                <div class="sf-side-col">
                    <div class="sf-side-card" style="background:#fff3cd; border-color:#ffeeba;">
                        <span style="font-size:12px; color:#856404; font-weight:600;">⚠ Verificación</span>
                        <p style="margin:4px 0 0 0; font-size:12px; color:#856404;">No encontramos duplicados potenciales de este Prospecto.</p>
                    </div>

                    <div class="sf-side-card">
                        <div style="font-size:13px; font-weight:700; margin-bottom:8px;">Vínculos rápidos</div>
                        <ul style="margin:0; padding-left:16px; font-size:12px; color:var(--sf-brand);">
                            <li style="margin-bottom:4px;"><a>Historial de aprobaciones (0)</a></li>
                            <li style="margin-bottom:4px;"><a>Archivos (0)</a></li>
                            <li><a>Notas (0)</a></li>
                        </ul>
                    </div>

                    <div class="sf-side-card">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span style="font-size:13px; font-weight:700;">Archivos (0)</span>
                            <button type="button" class="sf-btn-sub" style="padding:2px 8px; font-size:11px;">Cargar</button>
                        </div>
                        <div class="sf-drop-box">
                            <span style="font-size:12px; color:#514f4d;">Suelte archivos aquí</span>
                        </div>
                    </div>
                </div>

            </div>
        </form>
    </div>

</div>

<script>
    const tgMetadatos = {{ tg_meta_json|safe }};
    const HIERARCHY_DATA = {{ hierarchy_json|safe }};
    const modoLecturaGlobal = {% if modo_lectura %}true{% else %}false{% endif %};
    const unlockedIndexGlobal = {{ unlocked_idx }};
    
    // Función CRÍTICA para evitar que los campos ocultos bloqueen el HTML Form Submit
    function prepararFormularioParaEnvio() {
        document.querySelectorAll('[data-req="true"]').forEach(el => {
            el.removeAttribute('required'); // Quitamos required de todo el DOM al enviar
        });
        
        // Volvemos a colocar required solo en la pestaña que está visible actualmente
        const activeTg = document.getElementById('current_active_tg').value;
        const pantallaTarget = document.getElementById('pantalla-' + activeTg);
        if (pantallaTarget && !modoLecturaGlobal) {
            pantallaTarget.querySelectorAll('[data-req="true"]').forEach(el => {
                el.setAttribute('required', 'required');
            });
        }
    }

    function actualizarCascadaTG1(resetChildren) {
        const macroSelect = document.querySelector('select[name="macro_segmento"]');
        const subSelect = document.querySelector('select[name="sub_segmento"]');
        const actSelect = document.querySelector('select[name="actividad_economica"]');

        if (!macroSelect || !subSelect || !actSelect) return;

        const selMacro = macroSelect.value;
        const savedSub = subSelect.getAttribute('data-saved-val') || "";
        const savedAct = actSelect.getAttribute('data-saved-val') || "";

        const currentSub = resetChildren ? "" : (subSelect.value || savedSub);
        const currentAct = resetChildren ? "" : (actSelect.value || savedAct);

        if (selMacro && HIERARCHY_DATA[selMacro]) {
            const macroData = HIERARCHY_DATA[selMacro];
            
            // Sub-Segmentos
            if (!modoLecturaGlobal) subSelect.disabled = false;
            subSelect.innerHTML = '<option value="">--Seleccione Sub-Segmento--</option>';
            macroData.subs.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === currentSub) opt.selected = true;
                subSelect.appendChild(opt);
            });

            // Actividades Económicas
            let actList = macroData.acts;
            if (subSelect.value && macroData.sub_map[subSelect.value]) {
                actList = macroData.sub_map[subSelect.value];
            }

            if (!modoLecturaGlobal) actSelect.disabled = false;
            actSelect.innerHTML = '<option value="">--Seleccione Actividad Económica--</option>';
            actList.forEach(a => {
                const opt = document.createElement('option');
                opt.value = a;
                opt.textContent = a;
                if (a === currentAct) opt.selected = true;
                actSelect.appendChild(opt);
            });
        } else {
            subSelect.disabled = true;
            subSelect.innerHTML = '<option value="">--Seleccione primero Macro Segmento--</option>';
            actSelect.disabled = true;
            actSelect.innerHTML = '<option value="">--Seleccione primero Macro Segmento--</option>';
        }
    }

    function activarTollgate(tgId, tgIdx) {
        if (!modoLecturaGlobal && tgIdx !== undefined && tgIdx > unlockedIndexGlobal) {
            return false;
        }

        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.sf-chevron').forEach(el => el.classList.remove('active'));

        const pantallaTarget = document.getElementById('pantalla-' + tgId);
        const tabTarget = document.getElementById('tab-btn-' + tgId);

        if (pantallaTarget) pantallaTarget.style.display = 'block';
        if (tabTarget) tabTarget.classList.add('active');

        const activeElem = document.getElementById('current_active_tg');
        if (activeElem) activeElem.value = tgId;

        if (tgMetadatos[tgId]) {
            document.getElementById('header-objeto-fase').innerText = 'Objeto SF: ' + tgMetadatos[tgId].objeto + ' | Fase: ' + tgMetadatos[tgId].fase;
        }

        // Aplicamos la lógica de requirimientos
        prepararFormularioParaEnvio();

        if (tgId === 'TG1') {
            actualizarCascadaTG1(false);
        }
    }

    function actualizarHighlights() {
        const nombreElem = document.getElementById('input-nombre');
        const apellidosElem = document.getElementById('input-apellidos');
        const empresaElem = document.getElementById('input-empresa');
        const cargoElem = document.getElementById('input-cargo');
        const telefonoElem = document.getElementById('input-telefono');
        const emailElem = document.getElementById('input-email');

        const nombre = nombreElem ? nombreElem.value.trim() : '';
        const apellidos = apellidosElem ? apellidosElem.value.trim() : '';
        const empresa = empresaElem ? empresaElem.value.trim() : '';
        const cargo = cargoElem ? cargoElem.value.trim() : '';
        const telefono = telefonoElem ? telefonoElem.value.trim() : '';
        const email = emailElem ? emailElem.value.trim() : '';

        const nombreCompleto = (nombre || apellidos) ? (nombre + ' ' + apellidos) : '— Sin registrar —';
        document.getElementById('dyn-lead-title').innerText = nombreCompleto;
        document.getElementById('dyn-empresa').innerText = empresa || '—';
        document.getElementById('dyn-cargo').innerText = cargo || '—';
        document.getElementById('dyn-telefono').innerText = telefono || '—';
        document.getElementById('dyn-email').innerText = email || '—';
    }

    function volverALista() {
        document.getElementById('vista-detalle').style.display = 'none';
        document.getElementById('vista-lista').style.display = 'block';
    }

    // Inicializar estado
    document.addEventListener('DOMContentLoaded', function() {
        activarTollgate('{{ active_tg }}', {{ unlocked_idx }});
    });
</script>

</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    mostrar_detalle = False
    modo_lectura = False
    active_tg = 'TG0'
    unlocked_idx = 0
    prospecto_id = -1
    registro_actual = None
    tg_keys = list(TOLLGATES_DATA.keys())

    if request.method == 'POST':
        action_type = request.form.get('action_type', '')

        if action_type == 'ver_lectura':
            prospecto_id = int(request.form.get('prospecto_id', -1))
            if 0 <= prospecto_id < len(REGISTROS_PROSPECTOS):
                registro_actual = REGISTROS_PROSPECTOS[prospecto_id]
                unlocked_idx = registro_actual.get('unlocked_idx', 0)
                active_tg = registro_actual.get('active_tg', 'TG0')
                mostrar_detalle = True
                modo_lectura = True

        elif action_type == 'editar':
            prospecto_id = int(request.form.get('prospecto_id', -1))
            if 0 <= prospecto_id < len(REGISTROS_PROSPECTOS):
                registro_actual = REGISTROS_PROSPECTOS[prospecto_id]
                unlocked_idx = registro_actual.get('unlocked_idx', 0)
                active_tg = registro_actual.get('active_tg', 'TG0')
                mostrar_detalle = True
                modo_lectura = False

        elif action_type == 'cambiar_a_edicion':
            prospecto_id = int(request.form.get('prospecto_id', -1))
            if 0 <= prospecto_id < len(REGISTROS_PROSPECTOS):
                registro_actual = REGISTROS_PROSPECTOS[prospecto_id]
                unlocked_idx = registro_actual.get('unlocked_idx', 0)
                active_tg = registro_actual.get('active_tg', 'TG0')
                mostrar_detalle = True
                modo_lectura = False

        elif action_type == 'eliminar':
            prospecto_id = int(request.form.get('prospecto_id', -1))
            if 0 <= prospecto_id < len(REGISTROS_PROSPECTOS):
                eliminado = REGISTROS_PROSPECTOS.pop(prospecto_id)
                nom = (eliminado['datos'].get('nombre', '') + ' ' + eliminado['datos'].get('apellidos', '')).strip()
                mensaje = f"¡Prospecto '{nom or 'Seleccionado'}' eliminado correctamente!"
            mostrar_detalle = False

        elif action_type == 'nuevo':
            mostrar_detalle = True
            modo_lectura = False
            prospecto_id = -1
            unlocked_idx = 0
            active_tg = 'TG0'
            registro_actual = {"id": -1, "unlocked_idx": 0, "active_tg": "TG0", "datos": {}}

        elif action_type == 'avanzar':
            mostrar_detalle = True
            modo_lectura = False
            prospecto_id = int(request.form.get('prospecto_id', -1))
            current_active_tg = request.form.get('current_active_tg', 'TG0')
            unlocked_idx = int(request.form.get('unlocked_idx', 0))

            datos_capturados = {}
            for key, val in request.form.items():
                if key not in ['action_type', 'prospecto_id', 'current_active_tg', 'unlocked_idx']:
                    datos_capturados[key] = val.strip()

            if 0 <= prospecto_id < len(REGISTROS_PROSPECTOS):
                REGISTROS_PROSPECTOS[prospecto_id]['datos'].update(datos_capturados)
                registro_actual = REGISTROS_PROSPECTOS[prospecto_id]
            else:
                registro_actual = {
                    "id": len(REGISTROS_PROSPECTOS),
                    "unlocked_idx": unlocked_idx,
                    "active_tg": current_active_tg,
                    "datos": datos_capturados
                }
                REGISTROS_PROSPECTOS.append(registro_actual)
                prospecto_id = len(REGISTROS_PROSPECTOS) - 1

            nombre = datos_capturados.get('nombre', '')
            apellidos = datos_capturados.get('apellidos', '')
            empresa = datos_capturados.get('empresa', '')

            try:
                from office365.runtime.auth.user_credential import UserCredential
                from office365.sharepoint.client_context import ClientContext
                ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
                target_list = ctx.web.lists.get_by_title("BSV_Leads")
                
                target_list.add_item({
                    "Title": f"{nombre} {apellidos}".strip(),
                    "BSV_Empresa___Razon_Social__c": empresa,
                    "BSV_Cargo___Titulo__c": datos_capturados.get('cargo', ''),
                    "BSV_Email_Corporativo__c": datos_capturados.get('email', ''),
                    "BSV_Telefono_Contacto__c": datos_capturados.get('telefono', ''),
                    "BSV_WhatsApp__c": datos_capturados.get('whatsapp', ''),
                    "BSV_Pais___Region__c": datos_capturados.get('pais_region', ''),
                    "BSV_Macro_Segmento__c": datos_capturados.get('macro_segmento', ''),
                    "BSV_Sub_Segmento__c": datos_capturados.get('sub_segmento', ''),
                    "BSV_Actividad_Economica__c": datos_capturados.get('actividad_economica', ''),
                    "BSV_Tamano_Empresa__c": datos_capturados.get('tamano_empresa', '')
                })
                ctx.execute_query()
                # Éxito en SharePoint
                mensaje = f"¡Datos de {current_active_tg} guardados correctamente! Avanzando al siguiente Tollgate."
            except Exception as e:
                # Éxito Local
                mensaje = f"¡Datos de {current_active_tg} guardados correctamente! Avanzando al siguiente Tollgate."

            current_idx = tg_keys.index(current_active_tg) if current_active_tg in tg_keys else 0
            if current_idx < len(tg_keys) - 1:
                next_idx = current_idx + 1
                unlocked_idx = max(unlocked_idx, next_idx)
                active_tg = tg_keys[next_idx]
            else:
                active_tg = current_active_tg
                mostrar_detalle = False
                mensaje = f"¡Captura completada para {nombre} {apellidos}!"

            registro_actual['unlocked_idx'] = unlocked_idx
            registro_actual['active_tg'] = active_tg

        elif action_type == 'volver_lista':
            mostrar_detalle = False

    tg_meta_json = json.dumps({k: {"objeto": v["objeto"], "fase": v["fase"]} for k, v in TOLLGATES_DATA.items()})
    hierarchy_json = json.dumps(HIERARCHY_DATA, ensure_ascii=False)

    return render_template_string(
        HTML_TEMPLATE,
        tg_data=TOLLGATES_DATA,
        tg_keys=tg_keys,
        active_tg=active_tg,
        unlocked_idx=unlocked_idx,
        modo_lectura=modo_lectura,
        mostrar_detalle=mostrar_detalle,
        prospecto_id=prospecto_id,
        registro_actual=registro_actual,
        tg_meta_json=tg_meta_json,
        hierarchy_json=hierarchy_json,
        registros=REGISTROS_PROSPECTOS,
        mensaje=mensaje
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

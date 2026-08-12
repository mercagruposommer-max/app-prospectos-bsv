import os
import json
from datetime import datetime, timedelta
from flask import Flask, request, render_template_string

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- MATRIZ COMPLETA DE MACROSEGMENTOS (CARGADA DE FORMA SEGURA) ---
HIERARCHY_JSON_STR = r'''{"IIA — Infraestructura Inteligente y Automatización Industrial":{"subs":["IIA-01 — Robótica y Manufactura Avanzada","IIA-02 — Plataformas IIoT (Plataformas Industriales del Internet de las Cosas)","IIA-03 — Visión Artificial"],"acts":["333999 — Fabricación de otra maquinaria y equipo para la industria en general","334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico","333249 — Fabricación de maquinaria y equipo para otras industrias manufactureras"],"sub_map":{"IIA-01 — Robótica y Manufactura Avanzada":["333999 — Fabricación de otra maquinaria y equipo para la industria en general"],"IIA-02 — Plataformas IIoT (Plataformas Industriales del Internet de las Cosas)":["334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico"],"IIA-03 — Visión Artificial":["333249 — Fabricación de maquinaria y equipo para otras industrias manufactureras"]}},"EYU — Energía y Utilidades":{"subs":["EYU-01 — Equipo de Transmisión y Distribución de energía","EYU-02 — Equipo de Transmisión y Distribución de energía","EYU-03 — Generación y Distribución de Eneregía","EYU-04 — Agua","EYU-05 — Petróleo y Gas"],"acts":["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica","335311 — Fabricación de motores y generadores eléctricos","221111 — Generación de electricidad a partir de combustibles fósiles","221112 — Generación de electricidad a partir de energía hidráulica","221122 — Distribución de energía eléctrica","221119 — Generación de electricidad a partir de otro tipo de energía","221121 — Transmisión de energía eléctrica","221123 — Comercialización de energía eléctrica","237133 — Supervisión de construcción de obras de generación y conducción de energía eléctrica y de obras para telecomunicaciones","237131 — Construcción de obras de generación y conducción de energía eléctrica","221312 — Captación, tratamiento y suministro de agua (sector público)","237111 — Construcción de obras para el tratamiento, distribución y suministro de agua y drenaje","221311 — Captación, tratamiento y suministro de agua (sector privado)","213111 — Perforación de pozos petroleros y de gas","211111 — Extracción de petróleo y gas natural asociado","324110 — Refinación de petróleo","237122 — Construcción de plantas de refinería y petroquímica","221210 — Suministro de gas natural por ductos al consumidor final"],"sub_map":{"EYU-01 — Equipo de Transmisión y Distribución de energía":["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica"],"EYU-02 — Equipo de Transmisión y Distribución de energía":["335311 — Fabricación de motores y generadores eléctricos"],"EYU-03 — Generación y Distribución de Eneregía":["221111 — Generación de electricidad a partir de combustibles fósiles","221112 — Generación de electricidad a partir de energía hidráulica","221122 — Distribución de energía eléctrica","221119 — Generación de electricidad a partir de otro tipo de energía","221121 — Transmisión de energía eléctrica","221123 — Comercialización de energía eléctrica","237133 — Supervisión de construcción de obras de generación y conducción de energía eléctrica y de obras para telecomunicaciones","237131 — Construcción de obras de generación y conducción de energía eléctrica"],"EYU-04 — Agua":["221312 — Captación, tratamiento y suministro de agua (sector público)","237111 — Construcción de obras para el tratamiento, distribución y suministro de agua y drenaje","221311 — Captación, tratamiento y suministro de agua (sector privado)"],"EYU-05 — Petróleo y Gas":["213111 — Perforación de pozos petroleros y de gas","211111 — Extracción de petróleo y gas natural asociado","324110 — Refinación de petróleo","237122 — Construcción de plantas de refinería y petroquímica","221210 — Suministro de gas natural por ductos al consumidor final"]}},"ENR — Energías Renovables":{"subs":["ENR-01 — Energía Solar","ENR-02 — Energía Eólica","ENR-03 — Almacenamiento de energía y Baterías"],"acts":["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica","221113 — Generación de electricidad a partir de energía solar","333610 — Fabricación de motores de combustión interna, turbinas y transmisiones","221114 — Generación de electricidad a partir de energía eólica","335910 — Fabricación de acumuladores y pilas"],"sub_map":{"ENR-01 — Energía Solar":["335312 — Fabricación de equipo y aparatos de distribución de energía eléctrica","221113 — Generación de electricidad a partir de energía solar"],"ENR-02 — Energía Eólica":["333610 — Fabricación de motores de combustión interna, turbinas y transmisiones","221114 — Generación de electricidad a partir de energía eólica"],"ENR-03 — Almacenamiento de energía y Baterías":["335910 — Fabricación de acumuladores y pilas"]}},"EEE — Electrónica y Equipos Eléctricos":{"subs":["EEE-01 — Ensamble de PCB (Tarjetas de Circuitos)","EEE-03 — Equipo eléctrico que incluye Inversores y Convertidores de corriente","EEE-04 — Conectores","EEE-05 — Iluminación"],"acts":["334410 — Fabricación de componentes electrónicos","335999 — Fabricación de otros equipos eléctricos","335920 — Fabricación de cables de conducción eléctrica","335930 — Fabricación de enchufes, contactos, fusibles y accesorios","335132 — Fabricación de lámparas ornamentales","335131 — Fabricación de focos"],"sub_map":{"EEE-01 — Ensamble de PCB (Tarjetas de Circuitos)":["334410 — Fabricación de componentes electrónicos"],"EEE-03 — Equipo eléctrico que incluye Inversores y Convertidores de corriente":["335999 — Fabricación de otros equipos eléctricos"],"EEE-04 — Conectores":["335920 — Fabricación de cables de conducción eléctrica","335930 — Fabricación de enchufes, contactos, fusibles y accesorios"],"EEE-05 — Iluminación":["335132 — Fabricación de lámparas ornamentales","335131 — Fabricación de focos"]}},"AYD — Aeronáutica y Defensa":{"subs":["AYD-01 — Equipo Original de Aeronaves","AYD-02 — Motores Aeronáuticos","AYD-03 — Aviónica y Sistemas"],"acts":["336410 — Fabricación de equipo aeroespacial","333610 — Fabricación de motores de combustión interna, turbinas y transmisiones","334519 — Instrumentos de medición/control/navegación"],"sub_map":{"AYD-01 — Equipo Original de Aeronaves":["336410 — Fabricación de equipo aeroespacial"],"AYD-02 — Motores Aeronáuticos":["333610 — Fabricación de motores de combustión interna, turbinas y transmisiones"],"AYD-03 — Aviónica y Sistemas":["334519 — Instrumentos de medición/control/navegación"]}},"AYT — Automotriz y Transporte":{"subs":["AYT-01 — Motocicletas","AYT-02 — Automóviles y Camionetas","AYT-03 — Autobuses, Camiones y Remolques","AYT-04 — Otras Partes Automotrices","AYT-05 — Equipo Ferroviario"],"acts":["336991 — Fabricación de motocicletas","336110 — Fabricación de automóviles y camionetas","336310 — Motores de gasolina y sus partes","336120 — Fabricación de camiones y tractocamiones","336210 — Fabricación de carrocerías y remolques","336330 — Partes de sistemas de dirección y suspensión","336320 — Equipo eléctrico y electrónico y sus partes","336350 — Partes de sistemas de transmisión","336360 — Partes de asientos y accesorios interiores","326192 — Fabricación de autopartes de plástico con y sin reforzamiento","336390 — Otras partes para vehículos automotrices","336510 — Fabricación de equipo ferroviario"],"sub_map":{"AYT-01 — Motocicletas":["336991 — Fabricación de motocicletas"],"AYT-02 — Automóviles y Camionetas":["336110 — Fabricación de automóviles y camionetas","336310 — Motores de gasolina y sus partes"],"AYT-03 — Autobuses, Camiones y Remolques":["336120 — Fabricación de camiones y tractocamiones","336210 — Fabricación de carrocerías y remolques","336330 — Partes de sistemas de dirección y suspensión"],"AYT-04 — Otras Partes Automotrices":["336320 — Equipo eléctrico y electrónico y sus partes","336350 — Partes de sistemas de transmisión","336360 — Partes de asientos y accesorios interiores","326192 — Fabricación de autopartes de plástico con y sin reforzamiento","336390 — Otras partes para vehículos automotrices"],"AYT-05 — Equipo Ferroviario":["336510 — Fabricación de equipo ferroviario"]}},"MIP — Maquinaria Industrial y Equipo Pesado":{"subs":["MIP-01 — Maquinaria Industrial","MIP-02 — Maquinaria y Equipo para Construcción","MIP-03 — Otra Maquinaria","MIP-04 — Equipo para Levantar y trasladar","MIP-05 — Bombas y Compresores","MIP-06 — Refigeración industrial"],"acts":["333999 — Otra maquinaria y equipo para la industria en general","333610 — Motores de combustión interna, turbinas y transmisiones","333242 — Fabricación de maquinaria y equipo para la industria del hule y del plástico","333510 — Maquinaria para la industria metalmecánica","333130 — Maquinaria para la industria extractiva","333120 — Maquinaria y equipo para la construcción","333991 — Equipo para soldar y soldaduras","332410 — Fabricación de calderas industriales","333319 — Otra maquinaria y equipo para el comercio y los servicios","333920 — Maquinaria y equipo para levantar y trasladar","326220 — Bandas y mangueras de hule/plástico","333910 — Bombas y sistemas de bombeo","333412 — Equipo de refrigeración industrial y comercial"],"sub_map":{"MIP-01 — Maquinaria Industrial":["333999 — Otra maquinaria y equipo para la industria en general","333610 — Motores de combustión interna, turbinas y transmisiones","333242 — Fabricación de maquinaria y equipo para la industria del hule y del plástico","333510 — Maquinaria para la industria metalmecánica"],"MIP-02 — Maquinaria y Equipo para Construcción":["333130 — Maquinaria para la industria extractiva","333120 — Maquinaria y equipo para la construcción"],"MIP-03 — Otra Maquinaria":["333991 — Equipo para soldar y soldaduras","332410 — Fabricación de calderas industriales","333319 — Otra maquinaria y equipo para el comercio y los servicios"],"MIP-04 — Equipo para Levantar y trasladar":["333920 — Maquinaria y equipo para levantar y trasladar"],"MIP-05 — Bombas y Compresores":["326220 — Bandas y mangueras de hule/plástico","333910 — Bombas y sistemas de bombeo"],"MIP-06 — Refigeración industrial":["333412 — Equipo de refrigeración industrial y comercial"]}},"EAR — Equipo Agrícola, Pecuario, Sistemas de Riego y Jardinería":{"subs":["EAR-01 — Tractores, Cosechadoras Sembradoras y Maquinas de Jardinería","EAR-02 — Sistemas de Riego Agrícola","EAR-03 — Equipo pecuario"],"acts":["333111 — Maquinaria y equipo agrícola","237112 — Construcción de sistemas de riego agrícola","333112 — Maquinaria y equipo pecuario"],"sub_map":{"EAR-01 — Tractores, Cosechadoras Sembradoras y Maquinas de Jardinería":["333111 — Maquinaria y equipo agrícola"],"EAR-02 — Sistemas de Riego Agrícola":["237112 — Construcción de sistemas de riego agrícola"],"EAR-03 — Equipo pecuario":["333112 — Maquinaria y equipo pecuario"]}},"APG — Agricultura, Pesca, Gandería":{"subs":["APG-01 — Agricultura","APG-02 — Pesca"],"acts":["111160 — Cultivo de arroz","311921 — Beneficio del café","493130 — Almacenamiento de productos agrícolas que no requieren refrigeración","311311 — Elaboración de azúcar de caña","111219 — Cultivo de otras hortalizas","114111 — Pesca de camarón","114112 — Pesca de túnidos"],"sub_map":{"APG-01 — Agricultura":["111160 — Cultivo de arroz","311921 — Beneficio del café","493130 — Almacenamiento de productos agrícolas que no requieren refrigeración","311311 — Elaboración de azúcar de caña","111219 — Cultivo de otras hortalizas"],"APG-02 — Pesca":["114111 — Pesca de camarón","114112 — Pesca de túnidos"]}},"EYC — Edificación y Construcción":{"subs":["EYC-01 — Construcciones Metálicas","EYC-02 — Edificación","EYC-03 — Caminos, Puentes y Vías de Comunicación"],"acts":["332310 — Estructuras metálicas","236212 — Supervisión de edificación de naves y plantas industriales","236211 — Edificación de naves y plantas industriales","238122 — Montaje de estructuras de acero prefabricadas","339950 — Fabricación de anuncios y señalamientos","236221 — Edificación de inmuebles comerciales y de servicios","238222 — Instalaciones de sistemas centrales de aire acondicionado y calefacción","238210 — Instalaciones eléctricas en construcciones","238290 — Otras instalaciones y equipamiento en construcciones","238190 — Otros trabajos en exteriores","238311 — Colocación de muros falsos y aislamiento","236111 — Edificación de vivienda unifamiliar","236112 — Edificación de vivienda multifamiliar","238390 — Otros trabajos de acabados en edificaciones","238350 — Realización de trabajos de carpintería en el lugar de la construcción","237213 — Supervisión de división de terrenos y de construcción de obras de urbanización","237999 — Otras construcciones de ingeniería civil","238990 — Otros trabajos especializados para la construcción","237993 — Construcción de obras para transporte eléctrico y ferroviario","237311 — Instalación de señalamientos y protecciones en obras viales","237312 — Construcción de carreteras, puentes y similares"],"sub_map":{"EYC-01 — Construcciones Metálicas":["332310 — Estructuras metálicas","236212 — Supervisión de edificación de naves y plantas industriales","236211 — Edificación de naves y plantas industriales","238122 — Montaje de estructuras de acero prefabricadas","339950 — Fabricación de anuncios y señalamientos"],"EYC-02 — Edificación":["236221 — Edificación de inmuebles comerciales y de servicios","238222 — Instalaciones de sistemas centrales de aire acondicionado y calefacción","238210 — Instalaciones eléctricas en construcciones","238290 — Otras instalaciones y equipamiento en construcciones","238190 — Otros trabajos en exteriores","238311 — Colocación de muros falsos y aislamiento","236111 — Edificación de vivienda unifamiliar","236112 — Edificación de vivienda multifamiliar","238390 — Otros trabajos de acabados en edificaciones","238350 — Realización de trabajos de carpintería en el lugar de la construcción","237213 — Supervisión de división de terrenos y de construcción de obras de urbanización","237999 — Otras construcciones de ingeniería civil","238990 — Otros trabajos especializados para la construcción"],"EYC-03 — Caminos, Puentes y Vías de Comunicación":["237993 — Construcción de obras para transporte eléctrico y ferroviario","237311 — Instalación de señalamientos y protecciones en obras viales","237312 — Construcción de carreteras, puentes y similares"]}},"FME — Fabriciones Metálicas":{"subs":["FME-01 — Acero","FME-02 — Aluminio","FME-03 — Transformaciones Metálicas"],"acts":["331510 — Moldeo por fundición de piezas de hierro y acero","331111 — Complejos siderúrgicos","331112 — Fabricación de desbastes primarios y ferroaleaciones","331419 — Fundición y refinación de otros metales no ferrosos","331310 — Industria básica del aluminio","331520 — Moldeo por fundición de piezas metálicas no ferrosas","331210 — Fabricación de tubos y postes de hierro y acero","331220 — Fabricación de otros productos de hierro y acero","332320 — Productos de herrería","331412 — Fundición y refinación de metales preciosos","332420 — Tanques metálicos de calibre grueso","332810 — Recubrimientos y terminados metálicos","332212 — Utensilios de cocina metálicos","332110 — Forjados y troquelados","332430 — Envases metálicos de calibre ligero","332720 — Tornillos, tuercas, remaches y similares","332710 — Maquinado de piezas metálicas","332999 — Otros productos metálicos","332610 — Alambre y resortes","332910 — Válvulas metálicas","332991 — Baleros y rodamientos","331420 — Laminación secundaria de cobre","332510 — Fabricación de herrajes y cerraduras","331490 — Laminación secundaria de otros metales no ferrosos"],"sub_map":{"FME-01 — Acero":["331510 — Moldeo por fundición de piezas de hierro y acero","331111 — Complejos siderúrgicos","331112 — Fabricación de desbastes primarios y ferroaleaciones"],"FME-02 — Aluminio":["331419 — Fundición y refinación de otros metales no ferrosos","331310 — Industria básica del aluminio","331520 — Moldeo por fundición de piezas metálicas no ferrosas"],"FME-03 — Transformaciones Metálicas":["331210 — Fabricación de tubos y postes de hierro y acero","331220 — Fabricación de otros productos de hierro y acero","332320 — Productos de herrería","331412 — Fundición y refinación de metales preciosos","332420 — Tanques metálicos de calibre grueso","332810 — Recubrimientos y terminados metálicos","332212 — Utensilios de cocina metálicos","332110 — Forjados y troquelados","332430 — Envases metálicos de calibre ligero","332720 — Tornillos, tuercas, remaches y similares","332710 — Maquinado de piezas metálicas","332999 — Otros productos metálicos","332610 — Alambre y resortes","332910 — Válvulas metálicas","332991 — Baleros y rodamientos","331420 — Laminación secundaria de cobre","332510 — Fabricación de herrajes y cerraduras","331490 — Laminación secundaria de otros metales no ferrosos"]}},"PAM — Procesamiento de Alimentos y Manufactura Especializada":{"subs":["PAM-01 — Maquinaria para la industria alimentcicia","PAM-02 — Procesamiento de Alimentos","PAM-03 — Preservación de alimentos","PAM-04 — Empacado y Envasado de alimentos","PAM-05 — Molienda de Granos y Semillas","PAM-06 — Condimentos","PAM-07 — Dulces","PAM-08 — Bebidas y Tabaco","PAM-09 — Preparación de Alimentos"],"acts":["333992 — Maquinaria y equipo para envasar y empacar","333243 — Maquinaria y equipo para la industria alimentaria y de las bebidas","311110 — Elaboración de alimentos para animales","311230 — Elaboración de cereales para el desayuno","311923 — Elaboración de café instantáneo","311511 — Elaboración de leche líquida","311513 — Derivados y fermentos lácteos","311613 — Embutidos y conservas de carne","311811 — Panificación industrial","311820 — Galletas y pastas","311611 — Matanza de ganado, aves y otros","311999 — Otros alimentos","311910 — Botanas","311930 — Concentrados/jarabes/esencias","311830 — Tortillas de maíz y nixtamal","325130 — Pigmentos y colorantes sintéticos","311422 — Conservación de frutas y verduras","311991 — Gelatinas y otros postres en polvo","311412 — Congelación de guisos y otros","311411 — Congelación de frutas y verduras","311423 — Conservación por procesos distintos a la congelación","311612 — Corte y empacado de carne","311710 — Envasado de pescados y mariscos","311212 — Harina de trigo","311213 — Elaboración de harina de maíz","311221 — Féculas y almidones","311940 — Condimentos y aderezos","311340 — Dulces, chicles y confitería (no chocolate)","312111 — Refrescos y otras bebidas no alcohólicas","311222 — Aceites y grasas vegetales","312144 — Tequila","312120 — Cerveza","311993 — Elaboración de alimentos frescos para consumo inmediato"],"sub_map":{"PAM-01 — Maquinaria para la industria alimentcicia":["333992 — Maquinaria y equipo para envasar y empacar","333243 — Maquinaria y equipo para la industria alimentaria y de las bebidas"],"PAM-02 — Procesamiento de Alimentos":["311110 — Elaboración de alimentos para animales","311230 — Elaboración de cereales para el desayuno","311923 — Elaboración de café instantáneo","311511 — Elaboración de leche líquida","311513 — Derivados y fermentos lácteos","311613 — Embutidos y conservas de carne","311811 — Panificación industrial","311820 — Galletas y pastas","311611 — Matanza de ganado, aves y otros","311999 — Otros alimentos","311910 — Botanas","311930 — Concentrados/jarabes/esencias","311830 — Tortillas de maíz y nixtamal","325130 — Pigmentos y colorantes sintéticos","311422 — Conservación de frutas y verduras","311991 — Gelatinas y otros postres en polvo"],"PAM-03 — Preservación de alimentos":["311412 — Congelación de guisos y otros","311411 — Congelación de frutas y verduras","311423 — Conservación por procesos distintos a la congelación"],"PAM-04 — Empacado y Envasado de alimentos":["311612 — Corte y empacado de carne","311710 — Envasado de pescados y mariscos"],"PAM-05 — Molienda de Granos y Semillas":["311212 — Harina de trigo","311213 — Elaboración de harina de maíz","311221 — Féculas y almidones"],"PAM-06 — Condimentos":["311940 — Condimentos y aderezos"],"PAM-07 — Dulces":["311340 — Dulces, chicles y confitería (no chocolate)"],"PAM-08 — Bebidas y Tabaco":["312111 — Refrescos y otras bebidas no alcohólicas","311222 — Aceites y grasas vegetales","312144 — Tequila","312120 — Cerveza"],"PAM-09 — Preparación de Alimentos":["311993 — Elaboración de alimentos frescos para consumo inmediato"]}},"SYM — Salud y Medicina":{"subs":["SYM-01 — Dispositivos médicos","SYM-02 — Aparatos de diagnóstico","SYM-03 — Equipo de laboratorio","SYM-04 — Preparaciones farmacéuticas"],"acts":["334519 — Instrumentos de medición/control/medicina electrónica","339113 — Artículos oftálmicos","339111 — Equipo no electrónico para uso médico/dental/laboratorio","325412 — Preparaciones farmacéuticas"],"sub_map":{"SYM-01 — Dispositivos médicos":["334519 — Instrumentos de medición/control/medicina electrónica","339113 — Artículos oftálmicos"],"SYM-02 — Aparatos de diagnóstico":["334519 — Instrumentos de medición/control/medicina electrónica"],"SYM-03 — Equipo de laboratorio":["339111 — Equipo no electrónico para uso médico/dental/laboratorio"],"SYM-04 — Preparaciones farmacéuticas":["325412 — Preparaciones farmacéuticas"]}},"MRO — Mantenimiento MRO":{"subs":["MRO-01 (Materiales para Construcción - Cemento) — MRO · Manufactura de materiales e insumos","MRO-03 (Materiales para Aislamiento y Acabados) — MRO · Manufactura de materiales e insumos","MRO-04 (Materiales - Asfalto y Pavimentos) — MRO · Manufactura de materiales e insumos","MRO-05 (Agroquímicos) — MRO · Manufactura de materiales e insumos","MRO-06 (Químicos para el cuidado personal) — MRO · Manufactura de materiales e insumos","MRO-07 (Químicos básicos y Materiales Industriales) — MRO · Manufactura de materiales e insumos","MRO-08 — Automotriz y Reparación (Externos)","MRO-09 — Equipo de Protección Personal","MRO-010 — Detección de Fuego y Gas","MRO-11 — Reparación de Vehículos"],"acts":["327310 — Fabricación de cemento y productos a base de cemento en plantas integradas","327320 — Fabricación de concreto","327399 — Fabricación de otros productos de cemento y concreto","327330 — Fabricación de tubos y bloques de cemento y concreto","327391 — Fabricación de productos preesforzados de concreto","326198 — Fabricación de otros productos de plástico con reforzamiento","327420 — Fabricación de yeso y productos de yeso","327214 — Fabricación de fibra de vidrio","327999 — Fabricación de otros productos a base de minerales no metálicos","325211 — Fabricación de resinas sintéticas","327122 — Fabricación de azulejos y losetas no refractarias","324120 — Fabricación de productos de asfalto","325320 — Fabricación de pesticidas y otros agroquímicos, excepto fertilizantes","325310 — Fabricación de fertilizantes","325610 — Fabricación de jabones, limpiadores y dentífricos","325620 — Fabricación de cosméticos, perfumes y otras preparaciones de tocador","325999 — Fabricación de otros productos químicos","325180 — Fabricación de otros productos químicos básicos inorgánicos","325190 — Fabricación de otros productos químicos básicos orgánicos","325120 — Fabricación de gases industriales","325220 — Fabricación de fibras químicas","811312 — Reparación y mantenimiento de maquinaria y equipo industrial","811314 — Reparación y mantenimiento de maquinaria y equipo comercial y de servicios","811219 — Reparación y mantenimiento de otro equipo electrónico y de equipo de precisión","811313 — Reparación y mantenimiento de maquinaria y equipo para mover, levantar y acomodar materiales","811311 — Reparación y mantenimiento de maquinaria y equipo agropecuario y forestal","339111 — Fabricación de equipo no electrónico para uso médico, dental y para laboratorio","315999 — Confección de otros accesorios y prendas de vestir no clasificados en otra parte","315223 — Confección en serie de uniformes","334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico","811112 — Reparación del sistema eléctrico de automóviles y camiones","811111 — Reparación mecánica en general de automóviles y camiones","811115 — Reparación de suspensiones de automóviles y camiones","811113 — Rectificación de partes de motor de automóviles y camiones","811116 — Alineación y balanceo de automóviles y camiones","811129 — Instalación de cristales y otras reparaciones a la carrocería de automóviles y camiones","811119 — Otras reparaciones mecánicas de automóviles y camiones"],"sub_map":{"MRO-01 (Materiales para Construcción - Cemento) — MRO · Manufactura de materiales e insumos":["327310 — Fabricación de cemento y productos a base de cemento en plantas integradas","327320 — Fabricación de concreto","327399 — Fabricación de otros productos de cemento y concreto","327330 — Fabricación de tubos y bloques de cemento y concreto","327391 — Fabricación de productos preesforzados de concreto"],"MRO-03 (Materiales para Aislamiento y Acabados) — MRO · Manufactura de materiales e insumos":["326198 — Fabricación de otros productos de plástico con reforzamiento","327420 — Fabricación de yeso y productos de yeso","327214 — Fabricación de fibra de vidrio","327999 — Fabricación de otros productos a base de minerales no metálicos","325211 — Fabricación de resinas sintéticas","327122 — Fabricación de azulejos y losetas no refractarias"],"MRO-04 (Materiales - Asfalto y Pavimentos) — MRO · Manufactura de materiales e insumos":["324120 — Fabricación de productos de asfalto"],"MRO-05 (Agroquímicos) — MRO · Manufactura de materiales e insumos":["325320 — Fabricación de pesticidas y otros agroquímicos, excepto fertilizantes","325310 — Fabricación de fertilizantes"],"MRO-06 (Químicos para el cuidado personal) — MRO · Manufactura de materiales e insumos":["325610 — Fabricación de jabones, limpiadores y dentífricos","325620 — Fabricación de cosméticos, perfumes y otras preparaciones de tocador","325999 — Fabricación de otros productos químicos"],"MRO-07 (Químicos básicos y Materiales Industriales) — MRO · Manufactura de materiales e insumos":["325180 — Fabricación de otros productos químicos básicos inorgánicos","325190 — Fabricación de otros productos químicos básicos orgánicos","325120 — Fabricación de gases industriales","325220 — Fabricación de fibras químicas"],"MRO-08 — Automotriz y Reparación (Externos)":["811312 — Reparación y mantenimiento de maquinaria y equipo industrial","811314 — Reparación y mantenimiento de maquinaria y equipo comercial y de servicios","811219 — Reparación y mantenimiento de otro equipo electrónico y de equipo de precisión","811313 — Reparación y mantenimiento de maquinaria y equipo para mover, levantar y acomodar materiales","811311 — Reparación y mantenimiento de maquinaria y equipo agropecuario y forestal"],"MRO-09 — Equipo de Protección Personal":["339111 — Fabricación de equipo no electrónico para uso médico, dental y para laboratorio","315999 — Confección de otros accesorios y prendas de vestir no clasificados en otra parte","315223 — Confección en serie de uniformes"],"MRO-010 — Detección de Fuego y Gas":["334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico"],"MRO-11 — Reparación de Vehículos":["811112 — Reparación del sistema eléctrico de automóviles y camiones","811111 — Reparación mecánica en general de automóviles y camiones","811115 — Reparación de suspensiones de automóviles y camiones","811113 — Rectificación de partes de motor de automóviles y camiones","811116 — Alineación y balanceo de automóviles y camiones","811129 — Instalación de cristales y otras reparaciones a la carrocería de automóviles y camiones","811119 — Otras reparaciones mecánicas de automóviles y camiones"]}},"ELA — Electrodomésticos, Línea Blanca y Aires Acondicionados Compactos":{"subs":["ELA-01 — Electrodomésticos y Línea Blanca","ELA-02 — Aires Acondicionados Compactos"],"acts":["335220 — Fabricación de aparatos de línea blanca","333411 — Fabricación de equipo de aire acondicionado y calefacción"],"sub_map":{"ELA-01 — Electrodomésticos y Línea Blanca":["335220 — Fabricación de aparatos de línea blanca"],"ELA-02 — Aires Acondicionados Compactos":["333411 — Fabricación de equipo de aire acondicionado y calefacción"]}},"MUE — Muebles y Maderas":{"subs":["MUE-01 — Muebles del hogar","MUE-02 — Muebles de oficina","MUE-03 — Maderas"],"acts":["337120 — Fabricación de muebles, excepto cocinas integrales, muebles modulares de baño y muebles de oficina y estantería","337110 — Fabricación de cocinas integrales y muebles modulares de baño","327112 — Fabricación de muebles de baño","337210 — Fabricación de muebles de oficina y estantería","321920 — Fabricación de productos para embalaje y envases de madera","321993 — Fabricación de productos de madera de uso industrial","321210 — Fabricación de laminados y aglutinados de madera","321112 — Aserrado de tablas y tablones","321999 — Fabricación de otros productos de madera","321910 — Fabricación de productos de madera para la construcción"],"sub_map":{"MUE-01 — Muebles del hogar":["337120 — Fabricación de muebles, excepto cocinas integrales, muebles modulares de baño y muebles de oficina y estantería","337110 — Fabricación de cocinas integrales y muebles modulares de baño","327112 — Fabricación de muebles de baño"],"MUE-02 — Muebles de oficina":["337210 — Fabricación de muebles de oficina y estantería"],"MUE-03 — Maderas":["321920 — Fabricación de productos para embalaje y envases de madera","321993 — Fabricación de productos de madera de uso industrial","321210 — Fabricación de laminados y aglutinados de madera","321112 — Aserrado de tablas y tablones","321999 — Fabricación de otros productos de madera","321910 — Fabricación de productos de madera para la construcción"]}},"MAR — Marina y Construcción Naval":{"subs":["MAR-01 — Construcción Naval","MAR-02 — Motores marinos","MAR-03 — Sistemas de navegación"],"acts":["336610 — Fabricación de embarcaciones","333610 — Fabricación de motores de combustión interna, turbinas y transmisiones","334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico"],"sub_map":{"MAR-01 — Construcción Naval":["336610 — Fabricación de embarcaciones"],"MAR-02 — Motores marinos":["333610 — Fabricación de motores de combustión interna, turbinas y transmisiones"],"MAR-03 — Sistemas de navegación":["334519 — Fabricación de otros instrumentos de medición, control, navegación, y equipo médico electrónico"]}},"CME — Comercio al por menor":{"subs":["CME-01 — Tornillerías, Tlapalerías, Ferreterías y Refaccionarias","CME-02 — Supermercados y Tiendas de Autoservicio","CME-03 — Otros comercios","CME-04 — Carnes y Productos de Origen Animal"],"acts":["467111 — Comercio al por menor en ferreterías y tlapalerías","468211 — Comercio al por menor de partes y refacciones nuevas para automóviles, camionetas y camiones","468212 — Comercio al por menor de partes y refacciones usadas para automóviles, camionetas y camiones","462111 — Comercio al por menor en supermercados","467117 — Comercio al por menor de artículos para albercas y otros artículos","466314 — Comercio al por menor de lámparas ornamentales y candiles","466111 — Comercio al por menor de muebles para el hogar","465311 — Comercio al por menor de artículos de papelería","468420 — Comercio al por menor de aceites y grasas lubricantes, aditivos y similares para vehículos de motor","466112 — Comercio al por menor de electrodomésticos menores y aparatos de línea blanca","467114 — Comercio al por menor de vidrios y espejos","468411 — Comercio al por menor de gasolina y diésel","468412 — Comercio al por menor de gas L. P. en cilindros y para tanques estacionarios","434228 — Comercio al por mayor de ganado y aves en pie","461122 — Comercio al por menor de carne de aves","461121 — Comercio al por menor de carnes rojas"],"sub_map":{"CME-01 — Tornillerías, Tlapalerías, Ferreterías y Refaccionarias":["467111 — Comercio al por menor en ferreterías y tlapalerías","468211 — Comercio al por menor de partes y refacciones nuevas para automóviles, camionetas y camiones","468212 — Comercio al por menor de partes y refacciones usadas para automóviles, camionetas y camiones"],"CME-02 — Supermercados y Tiendas de Autoservicio":["462111 — Comercio al por menor en supermercados"],"CME-03 — Otros comercios":["467117 — Comercio al por menor de artículos para albercas y otros artículos","466314 — Comercio al por menor de lámparas ornamentales y candiles","466111 — Comercio al por menor de muebles para el hogar","465311 — Comercio al por menor de artículos de papelería","468420 — Comercio al por menor de aceites y grasas lubricantes, aditivos y similares para vehículos de motor","466112 — Comercio al por menor de electrodomésticos menores y aparatos de línea blanca","467114 — Comercio al por menor de vidrios y espejos","468411 — Comercio al por menor de gasolina y diésel","468412 — Comercio al por menor de gas L. P. en cilindros y para tanques estacionarios"],"CME-04 — Carnes y Productos de Origen Animal":["434228 — Comercio al por mayor de ganado y aves en pie","461122 — Comercio al por menor de carne de aves","461121 — Comercio al por menor de carnes rojas"]}},"CMA — Comercio al por mayor":{"subs":["CMA-01 — Materiales para construcción","CMA-02 — Maquinaria","CMA-03 — Materias Primas","CMA-04 — Empaque y Envase","CMA-05 — Fertilizantes","CMA-06 — Medicamentos y artículos farmaceúticos","CMA-07 — Automóviles, Camiones y Refacciones","CMA-08 — Alimentos","CMA-09 — Otros Comercios"],"acts":["434225 — Comercio al por mayor de equipo y material eléctrico","434221 — Comercio al por mayor de materiales metálicos para la construcción y la manufactura","434226 — Comercio al por mayor de pintura","434219 — Comercio al por mayor de otros materiales para la construcción, excepto de madera y metálicos","435210 — Comercio al por mayor de maquinaria y equipo para la construcción y la minería","434211 — Comercio al por mayor de cemento, tabique y grava","434224 — Comercio al por mayor de madera para la construcción y la industria","435319 — Comercio al por mayor de maquinaria y equipo para otros servicios y para actividades comerciales","435110 — Comercio al por mayor de maquinaria y equipo agropecuario, forestal y para la pesca","435419 — Comercio al por mayor de otra maquinaria y equipo de uso general","435220 — Comercio al por mayor de maquinaria y equipo para la industria manufacturera","434229 — Comercio al por mayor de otras materias primas para otras industrias","431180 — Comercio al por mayor de dulces y materias primas para repostería","431150 — Comercio al por mayor de semillas y granos alimenticios, especias y chiles secos","434223 — Comercio al por mayor de envases en general, papel y cartón para la industria","434111 — Comercio al por mayor de fertilizantes, plaguicidas y semillas para siembra","434112 — Comercio al por mayor de medicamentos veterinarios y alimentos para animales, excepto mascotas","433110 — Comercio al por mayor de productos farmacéuticos","434222 — Comercio al por mayor de productos químicos para la industria farmacéutica y para otro uso industrial","436111 — Comercio al por mayor de camiones","436112 — Comercio al por mayor de partes y refacciones nuevas para automóviles, camionetas y camiones","431130 — Comercio al por mayor de frutas y verduras frescas","431122 — Comercio al por mayor de carne de aves","431140 — Comercio al por mayor de huevo","431160 — Comercio al por mayor de leche y otros productos lácteos","431199 — Comercio al por mayor de otros alimentos","437312 — Intermediación de comercio al por mayor de productos para la industria, el comercio y los servicios","434311 — Comercio al por mayor de desechos metálicos","435412 — Comercio al por mayor de mobiliario y equipo de oficina","435411 — Comercio al por mayor de mobiliario, equipo, y accesorios de cómputo","435312 — Comercio al por mayor de artículos y accesorios para diseño y pintura artística","434230 — Comercio al por mayor de combustibles de uso industrial","432113 — Comercio al por mayor de cueros y pieles"],"sub_map":{"CMA-01 — Materiales para construcción":["434225 — Comercio al por mayor de equipo y material eléctrico","434221 — Comercio al por mayor de materiales metálicos para la construcción y la manufactura","434226 — Comercio al por mayor de pintura","434219 — Comercio al por mayor de otros materiales para la construcción, excepto de madera y metálicos","435210 — Comercio al por mayor de maquinaria y equipo para la construcción y la minería","434211 — Comercio al por mayor de cemento, tabique y grava","434224 — Comercio al por mayor de madera para la construcción y la industria"],"CMA-02 — Maquinaria":["435319 — Comercio al por mayor de maquinaria y equipo para otros servicios y para actividades comerciales","435110 — Comercio al por mayor de maquinaria y equipo agropecuario, forestal y para la pesca","435419 — Comercio al por mayor de otra maquinaria y equipo de uso general","435220 — Comercio al por mayor de maquinaria y equipo para la industria manufacturera"],"CMA-03 — Materias Primas":["434229 — Comercio al por mayor de otras materias primas para otras industrias","431180 — Comercio al por mayor de dulces y materias primas para repostería","431150 — Comercio al por mayor de semillas y granos alimenticios, especias y chiles secos"],"CMA-04 — Empaque y Envase":["434223 — Comercio al por mayor de envases en general, papel y cartón para la industria"],"CMA-05 — Fertilizantes":["434111 — Comercio al por mayor de fertilizantes, plaguicidas y semillas para siembra"],"CMA-06 — Medicamentos y artículos farmaceúticos":["434112 — Comercio al por mayor de medicamentos veterinarios y alimentos para animales, excepto mascotas","433110 — Comercio al por mayor de productos farmacéuticos","434222 — Comercio al por mayor de productos químicos para la industria farmacéutica y para otro uso industrial"],"CMA-07 — Automóviles, Camiones y Refacciones":["436111 — Comercio al por mayor de camiones","436112 — Comercio al por mayor de partes y refacciones nuevas para automóviles, camionetas y camiones"],"CMA-08 — Alimentos":["431130 — Comercio al por mayor de frutas y verduras frescas","431122 — Comercio al por mayor de carne de aves","431140 — Comercio al por mayor de huevo","431160 — Comercio al por mayor de leche y otros productos lácteos","431199 — Comercio al por mayor de otros alimentos"],"CMA-09 — Otros Comercios":["437312 — Intermediación de comercio al por mayor de productos para la industria, el comercio y los servicios","434311 — Comercio al por mayor de desechos metálicos","435412 — Comercio al por mayor de mobiliario y equipo de oficina","435411 — Comercio al por mayor de mobiliario, equipo, y accesorios de cómputo","435312 — Comercio al por mayor de artículos y accesorios para diseño y pintura artística","434230 — Comercio al por mayor de combustibles de uso industrial","432113 — Comercio al por mayor de cueros y pieles"]}},"MIN — Minería":{"subs":["MIN-01 — Minerales no Metálicos","MIN-02 — Minerales Metálicos"],"acts":["212321 — Minería de arena y grava para la construcción","212312 — Minería de mármol","212324 — Minería de sílice","327410 — Fabricación de cal","212311 — Minería de piedra caliza","327910 — Fabricación de productos abrasivos","212393 — Minería de barita","212398 — Minería de minerales no metálicos para productos químicos","212319 — Minería de otras piedras dimensionadas","212231 — Minería de cobre","212221 — Minería de oro","212222 — Minería de plata","212232 — Minería de plomo y zinc"],"sub_map":{"MIN-01 — Minerales no Metálicos":["212321 — Minería de arena y grava para la construcción","212312 — Minería de mármol","212324 — Minería de sílice","327410 — Fabricación de cal","212311 — Minería de piedra caliza","327910 — Fabricación de productos abrasivos","212393 — Minería de barita","212398 — Minería de minerales no metálicos para productos químicos","212319 — Minería de otras piedras dimensionadas"],"MIN-02 — Minerales Metálicos":["212231 — Minería de cobre","212221 — Minería de oro","212222 — Minería de plata","212232 — Minería de plomo y zinc"]}},"PYC — Papel, Cartón y productos derivados":{"subs":["PYC-01 — Cartón","PYC-04 — Papel","PYC-02 — Papel","PYC-03 — Otros Productos"],"acts":["322210 — Fabricación de envases de cartón","322132 — Fabricación de cartón y cartoncillo a partir de pulpa","322299 — Fabricación de otros productos de cartón y papel","322121 — Fabricación de papel en plantas integradas","322220 — Fabricación de bolsas de papel y productos celulósicos recubiertos y tratados","322122 — Fabricación de papel a partir de pulpa","322291 — Fabricación de pañales desechables y productos sanitarios"],"sub_map":{"PYC-01 — Cartón":["322210 — Fabricación de envases de cartón","322132 — Fabricación de cartón y cartoncillo a partir de pulpa","322299 — Fabricación de otros productos de cartón y papel"],"PYC-04 — Papel":["322121 — Fabricación de papel en plantas integradas","322220 — Fabricación de bolsas de papel y productos celulósicos recubiertos y tratados"],"PYC-02 — Papel":["322122 — Fabricación de papel a partir de pulpa"],"PYC-03 — Otros Productos":["322291 — Fabricación de pañales desechables y productos sanitarios"]}},"OIS — Otras Industrias y Servicios":{"subs":["OIS-01 — Otras Industrias Manufactureras","OIS-02 — Textiles","OIS-03 — Servicios Profesionales","OIS-04 — Impresión","OIS-05 — Transporte de personas","OIS-06 — Plástico y Hule","OIS-07 — Transporte de Mercancías","OIS-08 — Otros Servivios","OIS-09 — Vidrio","OIS-10 — Telecomunicaciones","OIS-11 — Residuos y Reciclaje"],"acts":["339999 — Otras industrias manufactureras","339930 — Fabricación de juguetes","339940 — Fabricación de artículos y accesorios para escritura, pintura, dibujo y actividades de oficina","325510 — Fabricación de pinturas y recubrimientos","339112 — Fabricación de material desechable de uso médico","313230 — Fabricación de telas no tejidas (comprimidas)","315229 — Confección en serie de otra ropa exterior de materiales textiles","313240 — Fabricación de telas de tejido de punto","315123 — Fabricación de ropa exterior de tejido de punto","313111 — Preparación e hilado de fibras duras naturales","313210 — Fabricación de telas anchas de tejido de trama","313320 — Fabricación de telas recubiertas","314992 — Fabricación de redes y otros productos de cordelería","541330 — Servicios de ingeniería","541810 — Agencias de publicidad","541510 — Servicios de diseño de sistemas de cómputo y servicios relacionados","561110 — Servicios de administración de negocios","541990 — Otros servicios profesionales, científicos y técnicos","541211 — Servicios de contabilidad y auditoría","813110 — Asociaciones, organizaciones y cámaras de productores, comerciantes y prestadores de servicios","551111 — Dirección y administración de grupos empresariales o corporativos","541110 — Bufetes jurídicos","323119 — Impresión de formas continuas y otros impresos","485111 — Transporte colectivo urbano y suburbano de pasajeros en autobuses de ruta fija","485115 — Transporte colectivo urbano y suburbano de pasajeros en autobuses que transitan en carril exclusivo","485210 — Transporte colectivo foráneo de pasajeros de ruta fija","326194 — Fabricación de otros productos de plástico de uso industrial sin reforzamiento","326193 — Fabricación de envases y contenedores de plástico para embalaje con y sin reforzamiento","326211 — Fabricación de llantas y cámaras","326120 — Fabricación de tubería y conexiones, y tubos para embalaje","326199 — Fabricación de otros productos de plástico sin reforzamiento","326150 — Fabricación de espumas y productos de uretano","326191 — Fabricación de productos de plástico para el hogar con y sin reforzamiento","326110 — Fabricación de bolsas y películas de plástico","325993 — Fabricación de resinas de plásticos reciclados","482110 — Transporte por ferrocarril","484129 — Otro autotransporte foráneo de carga general","488210 — Servicios relacionados con el transporte por ferrocarril","484232 — Autotransporte foráneo de materiales y residuos peligrosos","484121 — Autotransporte foráneo de productos agrícolas sin refrigeración","484239 — Otro autotransporte foráneo de carga especializado","488390 — Otros servicios relacionados con el transporte por agua","488320 — Servicios de carga y descarga para el transporte por agua","561990 — Otros servicios de apoyo a los negocios","561210 — Servicios integrales de apoyo a los negocios en instalaciones","541380 — Servicios de laboratorios de pruebas","111421 — Floricultura a cielo abierto","561730 — Servicios de instalación y mantenimiento de áreas verdes","532411 — Alquiler de maquinaria y equipo para construcción, minería y actividades forestales","488519 — Otros servicios de intermediación para el transporte de carga","561520 — Organización de excursiones y paquetes turísticos para agencias de viajes (Operadores de tours)","721111 — Hoteles con otros servicios integrados","812410 — Estacionamientos y pensiones para vehículos automotores","327219 — Fabricación de otros productos de vidrio","327213 — Fabricación de envases y ampolletas de vidrio","517312 — Operadores de servicios de telecomunicaciones inalámbricas","517910 — Otros servicios de telecomunicaciones","562121 — Recolección de residuos no peligrosos por el sector privado"],"sub_map":{"OIS-01 — Otras Industrias Manufactureras":["339999 — Otras industrias manufactureras","339930 — Fabricación de juguetes","339940 — Fabricación de artículos y accesorios para escritura, pintura, dibujo y actividades de oficina","325510 — Fabricación de pinturas y recubrimientos","339112 — Fabricación de material desechable de uso médico"],"OIS-02 — Textiles":["313230 — Fabricación de telas no tejidas (comprimidas)","315229 — Confección en serie de otra ropa exterior de materiales textiles","313240 — Fabricación de telas de tejido de punto","315123 — Fabricación de ropa exterior de tejido de punto","313111 — Preparación e hilado de fibras duras naturales","313210 — Fabricación de telas anchas de tejido de trama","313320 — Fabricación de telas recubiertas","314992 — Fabricación de redes y otros productos de cordelería"],"OIS-03 — Servicios Profesionales":["541330 — Servicios de ingeniería","541810 — Agencias de publicidad","541510 — Servicios de diseño de sistemas de cómputo y servicios relacionados","561110 — Servicios de administración de negocios","541990 — Otros servicios profesionales, científicos y técnicos","541211 — Servicios de contabilidad y auditoría","813110 — Asociaciones, organizaciones y cámaras de productores, comerciantes y prestadores de servicios","551111 — Dirección y administración de grupos empresariales o corporativos","541110 — Bufetes jurídicos"],"OIS-04 — Impresión":["323119 — Impresión de formas continuas y otros impresos"],"OIS-05 — Transporte de personas":["485111 — Transporte colectivo urbano y suburbano de pasajeros en autobuses de ruta fija","485115 — Transporte colectivo urbano y suburbano de pasajeros en autobuses que transitan en carril exclusivo","485210 — Transporte colectivo foráneo de pasajeros de ruta fija"],"OIS-06 — Plástico y Hule":["326194 — Fabricación de otros productos de plástico de uso industrial sin reforzamiento","326193 — Fabricación de envases y contenedores de plástico para embalaje con y sin reforzamiento","326211 — Fabricación de llantas y cámaras","326120 — Fabricación de tubería y conexiones, y tubos para embalaje","326199 — Fabricación de otros productos de plástico sin reforzamiento","326150 — Fabricación de espumas y productos de uretano","326191 — Fabricación de productos de plástico para el hogar con y sin reforzamiento","326110 — Fabricación de bolsas y películas de plástico","325993 — Fabricación de resinas de plásticos reciclados"],"OIS-07 — Transporte de Mercancías":["482110 — Transporte por ferrocarril","484129 — Otro autotransporte foráneo de carga general","488210 — Servicios relacionados con el transporte por ferrocarril","484232 — Autotransporte foráneo de materiales y residuos peligrosos","484121 — Autotransporte foráneo de productos agrícolas sin refrigeración","484239 — Otro autotransporte foráneo de carga especializado","488390 — Otros servicios relacionados con el transporte por agua","488320 — Servicios de carga y descarga para el transporte por agua"],"OIS-08 — Otros Servivios":["561990 — Otros servicios de apoyo a los negocios","561210 — Servicios integrales de apoyo a los negocios en instalaciones","541380 — Servicios de laboratorios de pruebas","111421 — Floricultura a cielo abierto","561730 — Servicios de instalación y mantenimiento de áreas verdes","532411 — Alquiler de maquinaria y equipo para construcción, minería y actividades forestales","488519 — Otros servicios de intermediación para el transporte de carga","561520 — Organización de excursiones y paquetes turísticos para agencias de viajes (Operadores de tours)","721111 — Hoteles con otros servicios integrados","812410 — Estacionamientos y pensiones para vehículos automotores"],"OIS-09 — Vidrio":["327219 — Fabricación de otros productos de vidrio","327213 — Fabricación de envases y ampolletas de vidrio"],"OIS-10 — Telecomunicaciones":["517312 — Operadores de servicios de telecomunicaciones inalámbricas","517910 — Otros servicios de telecomunicaciones"],"OIS-11 — Residuos y Reciclaje":["562121 — Recolección de residuos no peligrosos por el sector privado"]}}}'''

HIERARCHY_DATA = json.loads(HIERARCHY_JSON_STR)
MACRO_OPTS = list(HIERARCHY_DATA.keys())

# --- BASE DE DATOS LOCAL PROTOTIPO ---
REGISTROS_PROSPECTOS = []

# --- MATRIZ COMPLETA DE TOLLGATES (TG0 A TG13) ---
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
        {"id": "whatsapp", "campo": "WhatsApp (Número)", "tipo": "Teléfono", "req": False, "ayuda": "Número con formato internacional y código de país.", "notas": ""},
        {"id": "pais_region", "campo": "País / Región", "tipo": "Lista (picklist)", "req": True, "ayuda": "País o región donde opera el lead.", "notas": "", "opts": ["Estados Unidos", "Canadá", "México", "Centroamérica", "España"]}
      ]
    }, {
      "nombre": "Origen y Asignación",
      "campos": [
        {"id": "origen_prospecto", "campo": "Origen del Prospecto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indique el origen del prospecto.", "notas": "", "opts": ["Campaña de Marketing", "Web Lead", "Prospecto de Ventas", "Proyecto Especial (BSV)", "Otro"]},
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
        {"id": "tamano_empresa", "campo": "Tamaño de Empresa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estimación del tamaño de la empresa.", "notas": "", "opts": [
            "0 a 5 personas",
            "6 a 10 personas",
            "11 a 30 personas",
            "31 a 50 personas",
            "51 a 100 personas",
            "101 a 250 personas",
            "251 y más personas"
        ]},
        {"id": "banda", "campo": "Clasificación Preliminar", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación BSV preliminar del cliente.", "notas": "si es K o A, entonces es BSV - Normal", "opts": ["K", "A", "B", "C", "D"]}
      ]
    }]
  },
  "TG2": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Engagement y Señales",
      "campos": [
        {"id": "rol_contacto", "campo": "Rol del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rol del contacto en la decisión.", "notas": "", "opts": [
            "Decision Maker - CEO, CFO, Director de Planta (Plant Manager), Director de Cadena de Suministro (Supply Chain), Director de Manufactura.",
            "Influencer - Mandos Medios-Altos / Gerencia, Gerente de Producción, Gerente de Mantenimiento, Gerente de Proyecto (Project Manager).",
            "Evaluador Técnico - Compras y Líderes Técnicos, Gerente de Compras, Especialista de Compras (MRO), Ingeniero de Procesos, Especialista en Automatización/Robótica, Supervisor de Obra.",
            "Usuario final - Técnicos de mantenimiento, Inspectores de calidad en línea, Supervisores de Producción, Subcontratista."
        ]}
      ]
    }, {
      "nombre": "Enriquecimiento",
      "campos": [
        {"id": "num_plantas", "campo": "Número de Plantas / Ubicaciones", "tipo": "Número", "req": False, "ayuda": "Número estimado de plantas operativas.", "notas": ""},
        {"id": "contactos_adic", "campo": "Contactos Adicionales Identificados", "tipo": "UI_Contactos", "req": False, "ayuda": "Agregue uno o más contactos relacionados a esta cuenta.", "notas": ""}
      ]
    }]
  },
  "TG3": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Conversión MQL",
      "campos": [
        {"id": "codigo_xz", "campo": "Código XZ", "tipo": "Lista (picklist)", "req": True, "ayuda": "Código de clasificación XZ.", "notas": "Consumo", "opts": ["Consumo Directo (Manufactura - Construcción)", "Consumo Indirecto (Mantenimiento, Reparación y Operaciones)", "Consumo Mixto"]}
      ]
    }, {
      "nombre": "Notas de Transferencia",
      "campos": [
        {"id": "area_interes", "campo": "Área de interés", "tipo": "Checkboxes", "req": True, "ayuda": "Seleccione las áreas de interés del prospecto.", "notas": "", "opts": [
            "Tornillería",
            "Rodamientos",
            "Transmisiones de potencia mecánica",
            "Fabricaciones especiales (A medida / Grandes volúmenes)",
            "Consultoría de ingeniería (Optimización de líneas de Producción)"
        ]},
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
        {"id": "clasif_comercial", "campo": "Clasificación Comercial", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial del lead.", "notas": "", "opts": ["K", "A", "B", "C", "D"]},
        {"id": "justif_clasif", "campo": "Justificación de Clasificación", "tipo": "Texto largo", "req": True, "ayuda": "Evidencia de la lógica aplicada.", "notas": ""}
      ]
    }, {
      "nombre": "Asignación y SLA",
      "campos": [
        {"id": "osp_asignado", "campo": "Agente de Ventas Externo Asignado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Nombre del Agente de Ventas que recibe la cuenta.", "notas": "", "opts": []},
        {"id": "gerente_area", "campo": "Gerente de Área", "tipo": "Texto", "req": True, "ayuda": "Nombre del Gerente responsable (Auto-asignado por Geografía).", "notas": ""},
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
        {"id": "tipo_oportunidad", "campo": "Tipo de Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial de la oportunidad.", "notas": "", "opts": ["XR — Reactivación", "XP — Prospecto", "XS — Cross Sell Sommer"]},
        {"id": "aoi_estimado", "campo": "AOI Estimado (%)", "tipo": "Porcentaje (%)", "req": True, "ayuda": "¿Debe tener un porcentaje mínimo?", "notas": ""},
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
        {"id": "aoi_cuantificado", "campo": "AOI Cuantificado ($)", "tipo": "Moneda_Entero", "req": True, "ayuda": "Valor en enteros (sin decimales).", "notas": ""},
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
        {"id": "aoi_proyectado_ano1", "campo": "AOI Proyectado Año 1 ($)", "tipo": "Moneda_Entero", "req": True, "ayuda": "Valor en enteros (sin decimales).", "notas": ""},
        {"id": "aoi_proyectado_anos2_3", "campo": "AOI Proyectado Años 2-3 ($)", "tipo": "Moneda_Entero", "req": True, "ayuda": "Valor acumulado en enteros.", "notas": ""},
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

        /* ESTILOS DE PESTAÑAS Y BLOQUEO PROGRESIVO DE CHEVRONS */
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
        .sf-input:disabled, .sf-select:disabled, .sf-textarea:disabled, .sf-input[readonly] { background-color: #f3f3f3; color: #555555; cursor: not-allowed; border-color: #dddbda; pointer-events: none; }
        
        .sf-help-text { font-size: 11px; color: var(--sf-text-muted); margin-top: 5px; font-style: italic; }

        .sf-side-card { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; padding: 14px; }
        .sf-drop-box { border: 2px dashed var(--sf-border); border-radius: 4px; padding: 20px; text-align: center; background: #fafafa; margin-top: 8px; transition: background 0.2s; }
        .sf-drop-box:hover { background: #f0f0f0; }

        .alert-success { background-color: #d4edda; color: #155724; padding: 12px 16px; border-radius: 4px; margin: 12px; border: 1px solid #c3e6cb; font-size: 13px; font-weight:600; }
        
        .badge-read-only { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; }
        .badge-edit-mode { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; }

        /* CHECKBOX GROUP (TG3) */
        .sf-checkbox-group { display: flex; flex-direction: column; gap: 8px; }
        .sf-checkbox-label { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: normal; cursor: pointer; color: var(--sf-text-main); }
        .sf-checkbox-label input { width: 16px; height: 16px; cursor: pointer; }

        /* MODAL CONTACTOS ADICIONALES (Fichas) */
        .modal-overlay {
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5); z-index: 9999; align-items: center; justify-content: center;
        }
        .modal-content {
            background: #fff; padding: 20px; border-radius: 6px; width: 600px; max-width: 90%;
            max-height: 90vh; overflow-y: auto; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }
        .contact-card {
            border: 1px solid var(--sf-border); border-radius: 4px; padding: 10px;
            background: #fff; width: 100%; box-sizing: border-box; cursor: pointer;
            margin-bottom: 8px; transition: border-color 0.2s;
        }
        .contact-card:hover { border-color: var(--sf-brand); }
        .contact-card-header { font-weight: 700; color: var(--sf-brand); font-size: 13px; }
        .contact-card-body { display: none; margin-top: 8px; font-size: 12px; color: var(--sf-text-muted); padding-top: 8px; border-top: 1px dashed var(--sf-border); }
        .highlight-crm { border-left: 4px solid var(--sf-brand); }
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
                <!-- BOTÓN PARA CREAR NUEVO PROSPECTO -->
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
                                
                                <!-- CLIC EN EL NOMBRE: MODO LECTURA NO EDITABLE -->
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
                                            
                                            <!-- DEFAULT PREDETERMINADO EN MEXICO (TG0) -->
                                            {% if field['id'] == 'pais_region' and f_val == '' %}
                                                {% set f_val = 'México' %}
                                            {% endif %}

                                            <div class="sf-field-group" {% if field['tipo'] in ['Texto largo', 'UI_Contactos', 'Checkboxes'] %}style="grid-column: span 2;"{% endif %}>
                                                <div class="sf-label">
                                                    {% if field['req'] and not modo_lectura %}<span class="sf-req">*</span>{% endif %}{{ field['campo'] }}
                                                </div>

                                                {% if field['tipo'] == 'UI_Contactos' %}
                                                    <!-- INTERFAZ DINÁMICA DE CONTACTOS ADICIONALES -->
                                                    <div id="contactos-container" style="display:flex; flex-wrap:wrap; gap:10px; margin-bottom: 10px;">
                                                        <!-- Fichas de contactos se renderizan con JS -->
                                                    </div>
                                                    {% if not modo_lectura %}
                                                    <button type="button" class="sf-btn-sub" onclick="abrirModalContacto()">+ Agregar Contacto Adicional</button>
                                                    {% endif %}
                                                    <input type="hidden" name="{{ field['id'] }}" id="input_contactos_adic" value="{{ f_val }}">
                                                    
                                                {% elif field['tipo'] == 'Checkboxes' %}
                                                    <!-- INTERFAZ MULTIPLE CHECKBOX (EJ. AREAS DE INTERES) -->
                                                    <div class="sf-checkbox-group" data-req="{% if field['req'] %}true{% else %}false{% endif %}">
                                                        {% for opt in field['opts'] %}
                                                            <label class="sf-checkbox-label">
                                                                <input type="checkbox" name="{{ field['id'] }}" value="{{ opt }}" {% if opt in f_val %}checked{% endif %} {% if modo_lectura %}disabled{% endif %}>
                                                                {{ opt }}
                                                            </label>
                                                        {% endfor %}
                                                    </div>

                                                {% elif field['tipo'] == 'Lista (picklist)' %}
                                                    <select name="{{ field['id'] }}" 
                                                            class="sf-select"
                                                            data-saved-val="{{ f_val }}"
                                                            data-req="{% if field['req'] %}true{% else %}false{% endif %}"
                                                            {% if field['id'] == 'macro_segmento' %}onchange="actualizarCascadaTG1(true)"{% endif %}
                                                            {% if field['id'] == 'sub_segmento' %}onchange="actualizarCascadaTG1(false)"{% endif %}
                                                            {% if field['id'] == 'geografia' %}onchange="actualizarGeografia()"{% endif %}
                                                            {% if modo_lectura %}disabled{% endif %}>
                                                        <option value="">--Seleccione {{ field['campo'] }}--</option>
                                                        {% if field['opts'] %}
                                                            {% for opt in field['opts'] %}
                                                                {% set is_selected = (f_val == opt) or (field['id'] == 'pais_region' and opt == 'México' and f_val == '') %}
                                                                <option value="{{ opt }}" {% if is_selected %}selected{% endif %}>{{ opt }}</option>
                                                            {% endfor %}
                                                        {% endif %}
                                                    </select>
                                                {% elif field['tipo'] == 'Moneda_Entero' %}
                                                    <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" placeholder="Ej. 1,500,000" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="formatCurrency(this)">
                                                {% elif field['tipo'] == 'Texto largo' %}
                                                    <textarea name="{{ field['id'] }}" class="sf-textarea" rows="3" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %}>{{ f_val }}</textarea>
                                                {% elif field['tipo'] == 'Fecha' %}
                                                    <input type="date" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" min="{{ yesterday_str }}" {% if f_val %}readonly style="pointer-events:none; background-color:#f3f3f3;"{% endif %} {% if modo_lectura %}disabled{% endif %}>
                                                {% elif field['tipo'] == 'Email' %}
                                                    <input type="email" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] == 'Teléfono' %}
                                                    <input type="tel" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" placeholder="{% if field['id'] == 'whatsapp' %}+52 1 81 1234 5678{% else %}+52 81 0000 0000{% endif %}" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                                    <input type="number" step="any" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if modo_lectura %}disabled{% endif %}>
                                                {% else %}
                                                    <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" value="{{ f_val }}" class="sf-input" data-req="{% if field['req'] %}true{% else %}false{% endif %}" {% if field['id'] == 'gerente_area' %}readonly style="pointer-events:none; background-color:#f3f3f3;"{% endif %} {% if modo_lectura %}disabled{% endif %} oninput="actualizarHighlights()">
                                                {% endif %}

                                                <!-- DESCRIPCIÓN COLUMNA F -->
                                                <span class="sf-help-text">{{ field['ayuda'] or 'Capture la información solicitada en este campo.' }}</span>
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
                            <button type="submit" class="sf-btn-nuevo" id="btn-avanzar">
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
                            <!-- INPUT TIPO FILE OCULTO Y LIMITADO A FORMATOS -->
                            <input type="file" id="file_upload" accept=".pdf,.doc,.docx,.jpg,.jpeg,.png" style="display:none;" onchange="alert('Archivo ' + this.files[0].name + ' cargado exitosamente en el prototipo.');">
                            <button type="button" class="sf-btn-sub" style="padding:2px 8px; font-size:11px;" onclick="document.getElementById('file_upload').click()">Cargar</button>
                        </div>
                        <div class="sf-drop-box" onclick="document.getElementById('file_upload').click()" style="cursor:pointer;">
                            <span style="font-size:12px; color:#514f4d;">Suelte archivos aquí (.pdf, .doc, .jpg, .png)</span>
                        </div>
                    </div>
                </div>

            </div>
        </form>
    </div>

</div>

<!-- MODAL CONTACTOS ADICIONALES (FICHAS Y POPUP) -->
<div id="modal-contacto" class="modal-overlay">
    <div class="modal-content">
        <h3 style="margin-top:0; font-size:16px; color:var(--sf-brand);">Agregar Contacto Adicional</h3>
        <div style="font-size:12px; color:var(--sf-text-muted); margin-bottom:16px;">Registre la información de un nuevo contacto o asocie uno detectado en el CRM.</div>
        
        <div class="sf-field-grid" style="padding:0; grid-template-columns: 1fr 1fr;">
            <div class="sf-field-group">
                <label class="sf-label">Nombre</label>
                <input type="text" id="mod_nombre" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Apellidos</label>
                <input type="text" id="mod_apellidos" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Empresa</label>
                <input type="text" id="mod_empresa" class="sf-input" value="{{ d_curr.get('empresa', '') }}" disabled readonly>
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Cargo / Título</label>
                <input type="text" id="mod_cargo" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Email</label>
                <input type="email" id="mod_email" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Teléfono</label>
                <input type="text" id="mod_telefono" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">WhatsApp</label>
                <input type="text" id="mod_whatsapp" class="sf-input">
            </div>
            <div class="sf-field-group">
                <label class="sf-label">Rol del Contacto</label>
                <select id="mod_rol" class="sf-select">
                    <option value="">--Seleccione--</option>
                    <option value="Decision Maker">Decision Maker - CEO, CFO...</option>
                    <option value="Influencer">Influencer - Mandos Medios...</option>
                    <option value="Evaluador Técnico">Evaluador Técnico - Compras...</option>
                    <option value="Usuario final">Usuario final - Técnicos...</option>
                </select>
            </div>
        </div>
        
        <div style="text-align:right; margin-top:20px; border-top:1px solid var(--sf-border); padding-top:16px;">
            <button type="button" class="sf-btn-sub" onclick="cerrarModalContacto()" style="margin-right:8px;">Cancelar</button>
            <button type="button" class="sf-btn-nuevo" onclick="guardarModalContacto()">Guardar Contacto</button>
        </div>
    </div>
</div>

<script>
    const tgMetadatos = {{ tg_meta_json|safe }};
    const HIERARCHY_DATA = {{ hierarchy_json|safe }};
    const modoLecturaGlobal = {% if modo_lectura %}true{% else %}false{% endif %};
    const unlockedIndexGlobal = {{ unlocked_idx }};
    
    // --- LÓGICA DE ASIGNACIÓN GEOGRÁFICA DE GERENTES Y AGENTES (TG4) ---
    const GEO_MAPPING = {
        "Norte": { "agentes": ["Daniel Guadalupe Pérez Ontiveros", "Jesús Vázquez Campos"], "gerente": "Daniel Guadalupe Pérez Ontiveros" },
        "Centro": { "agentes": ["Octavio Velázquez Torres", "Roberto Alejandro Pallares Hernández", "Juan Carlos Resendiz Ríos"], "gerente": "Octavio Velázquez Torres" },
        "Occidente": { "agentes": ["Marcos Gustavo Díaz Ramírez", "Isaac Janin León Zárate"], "gerente": "Marcos Gustavo Díaz Ramírez" },
        "Bajio": { "agentes": ["José Fernando Merlos Salgado", "Miguel Ángel Pérez Córdova"], "gerente": "José Fernando Merlos Salgado" },
        "Golfo": { "agentes": ["Martín García Ramírez", "Erick Omar Benítez Flores", "Hernán Alfredo Mijangos Martínez"], "gerente": "Martín García Ramírez" }
    };

    function actualizarGeografia() {
        const geoSelect = document.querySelector('select[name="geografia"]');
        const agenteSelect = document.querySelector('select[name="osp_asignado"]');
        const gerenteInput = document.querySelector('input[name="gerente_area"]');

        if (!geoSelect || !agenteSelect || !gerenteInput) return;

        const selGeo = geoSelect.value;
        const savedAgente = agenteSelect.getAttribute('data-saved-val') || "";

        if (selGeo && GEO_MAPPING[selGeo]) {
            const map = GEO_MAPPING[selGeo];
            
            // Llenar Agentes
            if (!modoLecturaGlobal) agenteSelect.disabled = false;
            agenteSelect.innerHTML = '<option value="">--Seleccione--</option>';
            map.agentes.forEach(a => {
                const opt = document.createElement('option');
                opt.value = a;
                opt.textContent = a;
                if (a === savedAgente || a === agenteSelect.value) opt.selected = true;
                agenteSelect.appendChild(opt);
            });

            // Auto-llenar Gerente
            gerenteInput.value = map.gerente;

        } else {
            agenteSelect.innerHTML = '<option value="">--Seleccione Geografía en TG1--</option>';
            if (!modoLecturaGlobal) agenteSelect.disabled = true;
            gerenteInput.value = '';
        }
    }
    
    // --- LÓGICA DE FICHAS DE CONTACTOS ADICIONALES (POPUP Y CRM) ---
    let arrayContactos = [];

    function inicializarContactos() {
        const inp = document.getElementById('input_contactos_adic');
        if(inp && inp.value) {
            try {
                arrayContactos = JSON.parse(inp.value);
            } catch(e) {
                arrayContactos = [];
            }
        }
        
        const nombreEmpresa = document.getElementById('input-empresa');
        if(nombreEmpresa && nombreEmpresa.value.trim() !== '' && arrayContactos.length === 0) {
            arrayContactos.push({
                nombre: "Ana",
                apellidos: "García (CRM)",
                empresa: nombreEmpresa.value.trim(),
                cargo: "Gerente de Compras",
                email: "ana.garcia@crm.com",
                telefono: "+52 55 1234 5678",
                whatsapp: "",
                rol: "Evaluador Técnico",
                crm_badge: true
            });
            guardarEstadoContactos();
        }
        renderizarFichasContactos();
    }

    function abrirModalContacto() {
        const nombreEmpresa = document.getElementById('input-empresa');
        if (nombreEmpresa) {
            document.getElementById('mod_empresa').value = nombreEmpresa.value;
        }
        document.getElementById('modal-contacto').style.display = 'flex';
    }
    
    function cerrarModalContacto() {
        document.getElementById('modal-contacto').style.display = 'none';
        document.getElementById('mod_nombre').value = '';
        document.getElementById('mod_apellidos').value = '';
        document.getElementById('mod_cargo').value = '';
        document.getElementById('mod_email').value = '';
        document.getElementById('mod_telefono').value = '';
        document.getElementById('mod_whatsapp').value = '';
        document.getElementById('mod_rol').value = '';
    }
    
    function guardarModalContacto() {
        const c = {
            nombre: document.getElementById('mod_nombre').value,
            apellidos: document.getElementById('mod_apellidos').value,
            empresa: document.getElementById('mod_empresa').value,
            cargo: document.getElementById('mod_cargo').value,
            email: document.getElementById('mod_email').value,
            telefono: document.getElementById('mod_telefono').value,
            whatsapp: document.getElementById('mod_whatsapp').value,
            rol: document.getElementById('mod_rol').value,
            crm_badge: false
        };
        arrayContactos.push(c);
        guardarEstadoContactos();
        renderizarFichasContactos();
        cerrarModalContacto();
    }

    function guardarEstadoContactos() {
        const inp = document.getElementById('input_contactos_adic');
        if(inp) inp.value = JSON.stringify(arrayContactos);
    }

    function toggleFichaInfo(idx) {
        const body = document.getElementById('ficha-body-' + idx);
        body.style.display = body.style.display === 'none' ? 'block' : 'none';
    }

    function renderizarFichasContactos() {
        const container = document.getElementById('contactos-container');
        if(!container) return;
        container.innerHTML = '';
        
        arrayContactos.forEach((c, idx) => {
            let badgeHtml = c.crm_badge ? `<span style="background:#eef4fe; color:var(--sf-brand); padding:2px 6px; border-radius:4px; font-size:10px; margin-left:8px;">Contacto CRM Sugerido</span>` : '';
            let html = `
                <div class="contact-card ${c.crm_badge ? 'highlight-crm' : ''}" onclick="toggleFichaInfo(${idx})">
                    <div class="contact-card-header">
                        👤 ${c.nombre} ${c.apellidos} ${badgeHtml}
                    </div>
                    <div style="font-size:12px; color:#514f4d; margin-top:4px;">${c.cargo || 'Sin cargo'} - ${c.rol || 'Rol no definido'}</div>
                    <div class="contact-card-body" id="ficha-body-${idx}">
                        <strong>Empresa:</strong> ${c.empresa}<br>
                        <strong>Email:</strong> ${c.email}<br>
                        <strong>Teléfono:</strong> ${c.telefono}<br>
                        <strong>WhatsApp:</strong> ${c.whatsapp}
                    </div>
                </div>
            `;
            container.innerHTML += html;
        });
    }

    // --- FORMATEO DE MONEDA ABIERTA ---
    function formatCurrency(input) {
        let val = input.value.replace(/\D/g, ''); 
        if (val !== '') {
            val = parseInt(val, 10).toLocaleString('en-US'); 
        }
        input.value = val;
    }
    
    function prepararFormularioParaEnvio() {
        document.querySelectorAll('[data-req="true"]').forEach(el => {
            el.removeAttribute('required'); 
        });
        
        const activeTg = document.getElementById('current_active_tg').value;
        const pantallaTarget = document.getElementById('pantalla-' + activeTg);
        
        if (pantallaTarget && !modoLecturaGlobal) {
            pantallaTarget.querySelectorAll('[data-req="true"]').forEach(el => {
                // Evitamos ponerle required a los grupos de checkboxes si la validación HTML falla nativamente
                if(!el.classList.contains('sf-checkbox-group')) {
                    el.setAttribute('required', 'required');
                }
            });
            
            // Validación manual rápida para el grupo de Checkboxes (Áreas de interés TG3)
            const checkboxGroup = pantallaTarget.querySelector('.sf-checkbox-group[data-req="true"]');
            if (checkboxGroup) {
                const checkboxes = checkboxGroup.querySelectorAll('input[type="checkbox"]');
                let checked = false;
                checkboxes.forEach(cb => { if (cb.checked) checked = true; });
                if (!checked) {
                    alert("Debe seleccionar al menos un 'Área de interés'.");
                    throw new Error("Validation Failed");
                }
            }
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

        let currentSub = "";
        let currentAct = "";
        
        if (resetChildren) {
            currentSub = "";
            currentAct = "";
        } else {
            if (subSelect.options.length > 1) {
                currentSub = subSelect.value;
            } else {
                currentSub = savedSub;
            }
            if (actSelect.options.length > 1) {
                currentAct = actSelect.value;
            } else {
                currentAct = savedAct;
            }
        }

        if (selMacro && HIERARCHY_DATA[selMacro]) {
            const macroData = HIERARCHY_DATA[selMacro];
            
            if (!modoLecturaGlobal) subSelect.disabled = false;
            subSelect.innerHTML = '<option value="">--Seleccione Sub-Segmento--</option>';
            macroData.subs.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === currentSub) opt.selected = true;
                subSelect.appendChild(opt);
            });

            let actList = macroData.acts;
            if (currentSub && macroData.sub_map[currentSub]) {
                actList = macroData.sub_map[currentSub];
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
            
            subSelect.setAttribute('data-saved-val', currentSub);
            actSelect.setAttribute('data-saved-val', currentAct);
            
        } else {
            subSelect.disabled = true;
            subSelect.innerHTML = '<option value="">--Seleccione primero Macro Segmento--</option>';
            actSelect.disabled = true;
            actSelect.innerHTML = '<option value="">--Seleccione primero Macro Segmento--</option>';
            subSelect.setAttribute('data-saved-val', '');
            actSelect.setAttribute('data-saved-val', '');
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

        prepararFormularioParaEnvio();

        if (tgId === 'TG1') {
            actualizarCascadaTG1(false);
        }
        
        if (tgId === 'TG2') {
            inicializarContactos();
        }
        
        if (tgId === 'TG4') {
            actualizarGeografia();
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

    document.addEventListener('DOMContentLoaded', function() {
        activarTollgate('{{ active_tg }}', {{ unlocked_idx }});
        
        // Ejecutar formatCurrency() al cargar si hay campos con valores ya guardados
        document.querySelectorAll('input[oninput*="formatCurrency"]').forEach(el => {
            if (el.value) {
                formatCurrency(el);
            }
        });
        
        // Inicializar jerarquía si estamos en un prospecto cargado
        actualizarGeografia();
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
    
    # Calcular fecha mínima (Ayer) para bloqueo de inputs date
    yesterday_str = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

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
            for key in request.form.keys():
                if key not in ['action_type', 'prospecto_id', 'current_active_tg', 'unlocked_idx']:
                    vals = request.form.getlist(key)
                    if len(vals) > 1:
                        # Si es un grupo de checkboxes, agrupar con comas
                        datos_capturados[key] = ", ".join(v.strip() for v in vals)
                    else:
                        datos_capturados[key] = vals[0].strip()

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
                mensaje = f"¡Datos de {current_active_tg} guardados en SharePoint correctamente!"
            except Exception as e:
                mensaje = f"¡Datos de {current_active_tg} guardados correctamente! Avanzando al siguiente Tollgate."

            current_idx = tg_keys.index(current_active_tg) if current_active_tg in tg_keys else 0
            if current_idx < len(tg_keys) - 1:
                next_idx = current_idx + 1
                unlocked_idx = max(unlocked_idx, next_idx)
                active_tg = tg_keys[next_idx]
            else:
                active_tg = current_active_tg
                mostrar_detalle = False
                nombre_c = (datos_capturados.get('nombre', '') + " " + datos_capturados.get('apellidos', '')).strip()
                mensaje = f"¡Captura completa de los 14 Tollgates finalizada exitosamente para {nombre_c}!"

            registro_actual['unlocked_idx'] = unlocked_idx
            registro_actual['active_tg'] = active_tg

        elif action_type == 'volver_lista':
            mostrar_detalle = False

    tg_meta_json = json.dumps({k: {"objeto": v["objeto"], "fase": v["fase"]} for k, v in TOLLGATES_DATA.items()})

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
        hierarchy_json=HIERARCHY_JSON_STR,
        yesterday_str=yesterday_str,
        registros=REGISTROS_PROSPECTOS,
        mensaje=mensaje
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

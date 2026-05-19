# -*- coding: utf-8 -*-
import qrcode
import io
import random
import string
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def cargar_configuracion():
    config = {}
    nombre_archivo = "configuracion.txt"
    if not os.path.exists(nombre_archivo):
        with open(nombre_archivo, "w", encoding="utf-8") as f:
            f.write("URL_FORMULARIO=https://docs.google.com/forms/d/e/1FAIpQLSdfEXAMPLE/viewform\n")
            f.write("ID_PREGUNTA=entry.123456789\n")
            f.write("CANTIDAD_QRS=60\n")
        print(f"\n[!] Se creo '{nombre_archivo}'. Completalo con tus datos reales y vuelve a ejecutar.")
        return None
    
    with open(nombre_archivo, "r", encoding="utf-8") as f:
        for linea in f:
            if "=" in linea:
                clave, valor = linea.split("=", 1)
                config[clave.strip()] = valor.strip()
    return config

def generar_codigo_aleatorio(longitud=6):
    letras_y_numeros = string.ascii_uppercase + string.digits
    return ''.join(random.choice(letras_y_numeros) for _ in range(longitud))

def generar_sistema_maestro():
    config = cargar_configuracion()
    if not config: return

    url_base = config.get("URL_FORMULARIO")
    id_pregunta = config.get("ID_PREGUNTA")
    cantidad = int(config.get("CANTIDAD_QRS", 60))

    print("\n" + "="*60)
    print(" SISTEMA DE ASISTENCIA MAESTRO - REPOSITORIO OFICIAL")
    print("="*60)
    print(f"[+] Generando entorno automatizado para {cantidad} diapositivas dinámicas...")

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)

    registro_auditoria = ["=== LLAVE MAESTRA DE AUDITORIA DE TOKENS ==="]
    registro_auditoria.append("Cátedra de Clínica Psicológica y Psicoterapias - UNC.")
    registro_auditoria.append("Lista oficial de llaves hash criptográficas para auditoría de asistencia.\n")

    for i in range(1, cantidad + 1):
        codigo_secreto = generar_codigo_aleatorio()
        token = f"QR{i}_{codigo_secreto}"
        
        segundos_totales = i * 15
        minutos = segundos_totales // 60
        segundos = segundos_totales % 60
        
        registro_auditoria.append(f"Min. {minutos}:{segundos:02d}seg (Diapositiva {i}): {token}")
        
        link_final = f"{url_base}?usp=pp_url&{id_pregunta}={token}"
        
        qr = qrcode.QRCode(version=1, box_size=15, border=2)
        qr.add_data(link_final)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        image_stream = io.BytesIO()
        img.save(image_stream)
        image_stream.seek(0)

        slide = prs.slides.add_slide(prs.slide_layouts[6]) 

        txTitle = slide.shapes.add_textbox(Inches(0), Inches(0.3), Inches(13.33), Inches(1))
        tf_title = txTitle.text_frame
        tf_title.text = "Asistencia Clínica psicológica y psicoterapias"
        p_title = tf_title.paragraphs[0]
        p_title.font.name = 'Arial'
        p_title.font.bold, p_title.font.size = True, Pt(36)
        p_title.alignment = PP_ALIGN.CENTER

        qr_size = Inches(4.8)
        slide.shapes.add_picture(image_stream, (prs.slide_width-qr_size)/2, (prs.slide_height-qr_size)/2, height=qr_size)

        txFooter = slide.shapes.add_textbox(Inches(0), Inches(6.3), Inches(13.33), Inches(1))
        tf_footer = txFooter.text_frame
        tf_footer.text = "ESCANEE ESTE QR Y ESCRIBA LA PALABRA CLAVE"
        p_footer = tf_footer.paragraphs[0]
        p_footer.font.name = 'Arial'
        p_footer.font.bold, p_footer.font.size = True, Pt(32)
        p_footer.alignment = PP_ALIGN.CENTER

    prs.save("Asistencia_Maestra_Clinica.pptx")
    with open("Llave_Maestra_Clinica.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(registro_auditoria))
    
    print("="*60)
    print("[✅] PROCESO EXITOSO: 'Asistencia_Maestra_Clinica.pptx' y 'Llave_Maestra_Clinica.txt' creados.")
    print("="*60 + "\n")

if __name__ == "__main__":
    generar_sistema_maestro()
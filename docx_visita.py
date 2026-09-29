import io
from docx import Document
from docx.oxml.ns import qn


def _texto(valor) -> str:
    """Convierte a texto evitando 'nan' y el '.0' de números leídos como decimal."""
    if valor is None:
        return ""
    s = str(valor).strip()
    if s.lower() == "nan":
        return ""
    return s[:-2] if s.endswith(".0") else s


def _limpiar_celda(cell):
    """Deja la celda con un único párrafo: elimina tablas anidadas (la celda de
    Teléfono de la plantilla traía una, que la hacía más ancha) y párrafos extra."""
    tc = cell._tc
    for tbl in tc.findall(qn("w:tbl")):
        tc.remove(tbl)
    parrafos = tc.findall(qn("w:p"))
    for p in parrafos[1:]:
        tc.remove(p)
    if not parrafos:
        cell.add_paragraph()


def llenar_visita(template_path: str, datos, telefono: str) -> io.BytesIO:
    """Carga la plantilla de visita técnica y rellena la tabla con los datos del asegurado."""
    doc = Document(template_path)
    tabla = doc.tables[0]

    direccion = _texto(datos["Dirección Riesgo Asegurado"])
    comuna = _texto(datos.get("Comuna", ""))
    if comuna and comuna.lower() not in direccion.lower():
        direccion = f"{direccion}, {comuna}" if direccion else comuna

    mapping = {
        "Número siniestro":            _texto(datos["Num_Siniestro"]),
        "Numero carpeta":              _texto(datos["Nro_Carpeta"]),
        "Dirección Riesgo asegurado":  direccion,
        "Asegurado":                   _texto(datos["Asegurado"]),
        "RUT":                         _texto(datos["Rut"]),
        "Teléfono":                    telefono.strip(),
    }

    for row in tabla.rows:
        label = row.cells[0].text.strip()
        if label in mapping:
            cell = row.cells[1]
            _limpiar_celda(cell)
            p = cell.paragraphs[0]
            valor = mapping[label]
            if p.runs:
                p.runs[0].text = valor
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(valor)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

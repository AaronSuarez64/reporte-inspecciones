import io
import re
import requests
import pandas as pd

GRAPH_URL = "https://graph.microsoft.com/v1.0"


class OneDriveClient:
    """Cliente para Microsoft Graph API con refresco automático de token."""

    def __init__(self, client_id: str, client_secret: str, refresh_token: str,
                 authority: str = "common"):
        self._client_id     = client_id
        self._client_secret = client_secret
        self._refresh_token = refresh_token
        self._token_url     = f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/token"
        self._access_token  = None
        self._renovar_token()

    def _renovar_token(self):
        resp = requests.post(self._token_url, data={
            "client_id":     self._client_id,
            "client_secret": self._client_secret,
            "refresh_token": self._refresh_token,
            "grant_type":    "refresh_token",
        })
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(f"Error al renovar token: {data.get('error_description', data)}")
        self._access_token = data["access_token"]
        if "refresh_token" in data:
            self._refresh_token = data["refresh_token"]

    def _get(self, url: str, **kwargs) -> requests.Response:
        headers = {"Authorization": f"Bearer {self._access_token}"}
        resp = requests.get(url, headers=headers, **kwargs)
        if resp.status_code == 401:
            self._renovar_token()
            headers["Authorization"] = f"Bearer {self._access_token}"
            resp = requests.get(url, headers=headers, **kwargs)
        return resp

    def _listar_hijos(self, item_id: str) -> list:
        """Lista todos los elementos dentro de una carpeta por su ID."""
        url = f"{GRAPH_URL}/me/drive/items/{item_id}/children"
        items = []
        while url:
            data = self._get(url).json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def descargar_por_id(self, item_id: str) -> io.BytesIO:
        """Descarga un archivo por su ID de OneDrive."""
        url  = f"{GRAPH_URL}/me/drive/items/{item_id}/content"
        resp = self._get(url)
        resp.raise_for_status()
        return io.BytesIO(resp.content)

    def leer_excel(self, item_id: str, sheet_name: str) -> pd.DataFrame:
        buf = self.descargar_por_id(item_id)
        return pd.read_excel(buf, sheet_name=sheet_name)

    def obtener_imagenes(self, fotos_item_id: str, nro_carpeta) -> tuple:
        """
        Navega fotos_item_id → carpeta que empieza con nro_carpeta → Fotos y Videos.
        Devuelve (lista_imagenes, mensaje_error_o_None).
        """
        # Buscar carpeta del caso
        carpetas = self._listar_hijos(fotos_item_id)
        carpeta = next(
            (c for c in carpetas
             if c.get("folder") and c["name"].startswith(str(nro_carpeta))),
            None,
        )
        if not carpeta:
            return [], f"No se encontró carpeta para el número {nro_carpeta}."

        # Buscar subcarpeta "Fotos y Videos"
        hijos = self._listar_hijos(carpeta["id"])
        sub = next(
            (h for h in hijos if h.get("folder") and h["name"] == "Fotos y Videos"),
            None,
        )
        if not sub:
            return [], "No se encontró la subcarpeta 'Fotos y Videos'."

        # Listar y ordenar imágenes
        archivos = self._listar_hijos(sub["id"])
        imagenes = []
        for f in archivos:
            nombre = f["name"].lower()
            if nombre.endswith((".jpg", ".jpeg", ".png", ".heic")):
                match = re.match(r"(.+?)_(\d+(?:\.\d+)?)\.(jpg|jpeg|png|heic)", nombre)
                if match:
                    imagenes.append({
                        "descripcion": match.group(1),
                        "posicion":    float(match.group(2)),
                        "item_id":     f["id"],
                        "extension":   match.group(3),
                    })
        imagenes.sort(key=lambda x: x["posicion"])
        return imagenes, None

import pandas as pd
import requests
from typing import Optional, Tuple

# Domains d'email grand public (souvent pas un domaine de site client)
GENERIC_EMAIL_DOMAINS = {
    "gmail.com", "outlook.com", "hotmail.com", "live.com",
    "yahoo.com", "icloud.com", "proton.me", "protonmail.com",
    "orange.fr", "free.fr", "skynet.be", "telenet.be", "sfr.fr",
    "gmx.com", "gmx.net"
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; CMSDetectorBot/1.0)"
}


def extract_domain_from_email(email: str) -> Optional[str]:
    """contact@monsite.com -> monsite.com"""
    if not isinstance(email, str):
        return None
    email = email.strip()
    if "@" not in email:
        return None

    parts = email.split("@")
    if len(parts) != 2:
        return None

    domain = parts[1].strip().lower()
    if not domain or domain in GENERIC_EMAIL_DOMAINS:
        return None

    # nettoyage basique
    domain = domain.replace(">", "").replace("<", "").replace(";", "").replace(",", "")
    return domain if "." in domain else None


def guess_website_url(domain: str) -> Tuple[Optional[str], Optional[requests.Response]]:
    """
    Essaie :
    https://domain, https://www.domain, http://domain, http://www.domain
    Retourne (url_finale, response) si un code 200-399 répond, sinon (None, None)
    """
    if not domain:
        return None, None

    schemes = ["https://", "http://"]
    prefixes = ["", "www."]

    for scheme in schemes:
        for prefix in prefixes:
            url = f"{scheme}{prefix}{domain}"
            try:
                resp = requests.get(
                    url,
                    headers=HEADERS,
                    timeout=10,
                    allow_redirects=True
                )
                if 200 <= resp.status_code < 400 and (resp.text is not None):
                    return resp.url, resp
            except requests.RequestException:
                continue

    return None, None


def detect_cms_from_response(resp: Optional[requests.Response]) -> str:
    """Détecte le CMS / e-shop via HTML + headers"""
    if resp is None:
        return "Unknown / No site"

    html = (resp.text or "").lower()
    headers = {k.lower(): str(v).lower() for k, v in (resp.headers or {}).items()}
    url = (resp.url or "").lower()

    # Shopify
    if (
        "cdn.shopify.com" in html
        or "shopify-checkout" in html
        or "myshopify.com" in url
        or "x-shopify-stage" in headers
    ):
        return "Shopify"

    # Wix
    if (
        "wixstatic.com" in html
        or "wix.com" in html
        or "x-wix-request-id" in headers
    ):
        return "Wix"

    # Odoo (website/ecommerce)
    if (
        "web.assets_frontend" in html
        or "/web/content/" in html
        or 'meta name="generator" content="odoo"' in html
        or "odoo" in headers.get("set-cookie", "")
    ):
        return "Odoo"

    # WordPress / WooCommerce
    if (
        "/wp-content/" in html
        or "/wp-includes/" in html
        or "wp-json" in html
        or 'meta name="generator" content="wordpress' in html
    ):
        if "woocommerce" in html:
            return "WordPress + WooCommerce"
        return "WordPress"

    # PrestaShop
    if "prestashop" in html or "powered by prestashop" in html:
        return "PrestaShop"

    # Squarespace
    if "static1.squarespace.com" in html or "squarespace.com" in html:
        return "Squarespace"

    return "Unknown / Custom"


def find_email_column(df: pd.DataFrame) -> str:
    """
    Essaie de trouver la colonne email :
    1) colonne nommée 'email'
    2) sinon, colonne qui contient le + de valeurs avec '@'
    3) sinon, première colonne
    """
    # 1) nom exact
    for col in df.columns:
        if str(col).strip().lower() == "email":
            return col

    # 2) heuristique: max de '@'
    best_col = None
    best_score = -1

    for col in df.columns:
        try:
            series = df[col].astype(str)
            score = series.str.contains("@", na=False).sum()
            if score > best_score:
                best_score = score
                best_col = col
        except Exception:
            continue

    if best_col is not None and best_score > 0:
        return best_col

    # 3) fallback
    return df.columns[0]


def detect_cms_for_email_with_url(email: object) -> Tuple[str, str]:
    """Retourne (cms, url_detected) en sécurisant NaN / types pandas."""
    email_str = "" if pd.isna(email) else str(email).strip()

    domain = extract_domain_from_email(email_str)
    if domain is None:
        return "No domain / Generic email", ""

    url, resp = guess_website_url(domain)
    if url is None or resp is None:
        return "Unknown / No site", ""

    cms = detect_cms_from_response(resp)
    return cms, url


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajoute:
    - cms_detected
    - detected_url
    - error_detail (raison réelle si Error)
    """
    email_col = find_email_column(df)

    cms_results = []
    urls_results = []
    error_details = []

    for email in df[email_col]:
        try:
            cms, final_url = detect_cms_for_email_with_url(email)
            cms_results.append(cms)
            urls_results.append(final_url)
            error_details.append("")
        except Exception as e:
            cms_results.append("Error")
            urls_results.append("")
            error_details.append(f"{type(e).__name__}: {e}")

    df_out = df.copy()
    df_out["cms_detected"] = cms_results
    df_out["detected_url"] = urls_results
    df_out["error_detail"] = error_details
    return df_out


def process_excel(input_path: str, output_path: Optional[str] = None) -> str:
    """Mode fichier->fichier (optionnel)."""
    df = pd.read_excel(input_path)
    df_out = process_dataframe(df)

    if output_path is None:
        if input_path.lower().endswith(".xlsx"):
            output_path = input_path[:-5] + "_with_cms.xlsx"
        else:
            output_path = input_path + "_with_cms.xlsx"

    df_out.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


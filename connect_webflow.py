#!/usr/bin/env python3
"""
Conecta ao site Indicium AI via Webflow API e atualiza posts PT-BR
que mencionam 'Indicium' sem 'Indicium AI' (rebranding).

Uso:
  export WEBFLOW_API_TOKEN=<seu_token>
  python connect_webflow.py              # modo dry-run (apenas inspeciona)
  python connect_webflow.py --update     # aplica as atualizações
  python connect_webflow.py --publish    # aplica e publica as alterações
"""

import json
import os
import sys
import time
import argparse
import re
import requests

WEBFLOW_API_BASE = "https://api.webflow.com/v2"
MAPPING_FILE = "indicium-blog-posts-mapping.json"

SITE_ID = "695fda9a67d79f8741c387ba"
COLLECTION_ID = "69666cc2070439f1bc1173cf"
LOCALE = "pt-BR"

# Regex: "Indicium" não seguido de " AI" (case-sensitive)
PATTERN_OLD = re.compile(r"Indicium(?! AI)")
REPLACEMENT = "Indicium AI"


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _get(token: str, path: str, params: dict | None = None) -> dict:
    url = f"{WEBFLOW_API_BASE}{path}"
    resp = requests.get(url, headers=_headers(token), params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _patch(token: str, path: str, body: dict, params: dict | None = None) -> dict:
    url = f"{WEBFLOW_API_BASE}{path}"
    resp = requests.patch(url, headers=_headers(token), json=body, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _post(token: str, path: str, body: dict) -> dict:
    url = f"{WEBFLOW_API_BASE}{path}"
    resp = requests.post(url, headers=_headers(token), json=body, timeout=30)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# Webflow API calls
# ---------------------------------------------------------------------------

def get_site(token: str) -> dict:
    return _get(token, f"/sites/{SITE_ID}")


def get_collection(token: str) -> dict:
    return _get(token, f"/collections/{COLLECTION_ID}")


def get_item(token: str, item_id: str) -> dict:
    return _get(token, f"/collections/{COLLECTION_ID}/items/{item_id}", params={"locale": LOCALE})


def patch_item(token: str, item_id: str, field_data: dict) -> dict:
    return _patch(
        token,
        f"/collections/{COLLECTION_ID}/items/{item_id}",
        body={"fieldData": field_data},
        params={"locale": LOCALE},
    )


def publish_items(token: str, item_ids: list[str]) -> dict:
    return _post(
        token,
        f"/collections/{COLLECTION_ID}/items/publish",
        body={"itemIds": item_ids},
    )


# ---------------------------------------------------------------------------
# Business logic
# ---------------------------------------------------------------------------

def load_mapping() -> dict:
    with open(MAPPING_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def find_fields_with_old_name(field_data: dict) -> dict[str, str]:
    """Retorna apenas os campos de texto que contêm 'Indicium' sem ' AI'."""
    dirty = {}
    for key, value in field_data.items():
        if isinstance(value, str) and PATTERN_OLD.search(value):
            dirty[key] = value
    return dirty


def apply_replacement(text: str) -> str:
    return PATTERN_OLD.sub(REPLACEMENT, text)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Webflow Indicium AI branding fixer")
    parser.add_argument("--update", action="store_true", help="Aplica as correções nos itens")
    parser.add_argument("--publish", action="store_true", help="Aplica e publica as correções")
    args = parser.parse_args()

    token = os.environ.get("WEBFLOW_API_TOKEN")
    if not token:
        print("Erro: variável de ambiente WEBFLOW_API_TOKEN não definida.", file=sys.stderr)
        sys.exit(1)

    do_update = args.update or args.publish
    do_publish = args.publish

    # ------------------------------------------------------------------
    # 1. Verificar conexão com o site
    # ------------------------------------------------------------------
    print("=" * 60)
    print("Conectando ao site Indicium AI via Webflow API...")
    try:
        site = get_site(token)
    except requests.HTTPError as e:
        print(f"Erro ao acessar o site: {e}", file=sys.stderr)
        sys.exit(1)

    site_name = site.get("displayName") or site.get("name") or SITE_ID
    print(f"  Site: {site_name} ({SITE_ID})")

    collection = get_collection(token)
    col_name = collection.get("displayName") or collection.get("name") or COLLECTION_ID
    print(f"  Collection: {col_name} ({COLLECTION_ID})")

    # ------------------------------------------------------------------
    # 2. Carregar mapeamento
    # ------------------------------------------------------------------
    mapping = load_mapping()
    posts = mapping["posts"]
    print(f"\nMapeamento carregado: {len(posts)} posts PT-BR para verificar.")

    # ------------------------------------------------------------------
    # 3. Inspecionar / atualizar cada post
    # ------------------------------------------------------------------
    print(f"\n{'='*60}")
    mode_label = "ATUALIZANDO" if do_update else "INSPECIONANDO (dry-run)"
    print(f"{mode_label} — locale: {LOCALE}\n")

    updated_ids: list[str] = []
    errors: list[str] = []

    for i, post in enumerate(posts, 1):
        item_id = post["id"]
        print(f"[{i:02d}/{len(posts)}] {post['name']}")

        try:
            item = get_item(token, item_id)
        except requests.HTTPError as e:
            print(f"       !! Erro ao buscar item: {e}")
            errors.append(item_id)
            continue

        field_data = item.get("fieldData", {})
        dirty_fields = find_fields_with_old_name(field_data)

        if not dirty_fields:
            print("       OK — nenhum campo com 'Indicium' sem 'AI'.")
            continue

        for field, old_value in dirty_fields.items():
            snippet = old_value[:80].replace("\n", " ")
            print(f"       Campo '{field}': …{snippet}…")

        if do_update:
            new_field_data = {k: apply_replacement(v) for k, v in dirty_fields.items()}
            try:
                patch_item(token, item_id, new_field_data)
                print(f"       -> Atualizado ({len(dirty_fields)} campo(s)).")
                updated_ids.append(item_id)
                time.sleep(0.2)  # respeitar rate limits
            except requests.HTTPError as e:
                print(f"       !! Erro ao atualizar: {e}")
                errors.append(item_id)
        else:
            print(f"       -> Precisa atualizar {len(dirty_fields)} campo(s). (use --update para aplicar)")

    # ------------------------------------------------------------------
    # 4. Publicar (opcional)
    # ------------------------------------------------------------------
    if do_publish and updated_ids:
        print(f"\n{'='*60}")
        print(f"Publicando {len(updated_ids)} itens atualizados...")
        try:
            result = publish_items(token, updated_ids)
            published = result.get("publishedItemIds", updated_ids)
            print(f"  Publicados: {len(published)} itens.")
        except requests.HTTPError as e:
            print(f"  !! Erro ao publicar: {e}", file=sys.stderr)

    # ------------------------------------------------------------------
    # 5. Resumo
    # ------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("RESUMO")
    print(f"  Posts verificados : {len(posts)}")
    if do_update:
        print(f"  Posts atualizados : {len(updated_ids)}")
    if errors:
        print(f"  Erros             : {len(errors)} ({', '.join(errors)})")
    if not do_update:
        print("\n  Execute com --update para aplicar as correções.")
        print("  Execute com --publish para aplicar e publicar automaticamente.")


if __name__ == "__main__":
    main()

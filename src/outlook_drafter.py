"""Read inbox, identify emails needing a response, create Outlook drafts. NEVER sends."""

import json
import os
import re
import anthropic
from .graph_client import GraphClient

TRIAGE_PROMPT = """Tu es l'assistant email de {user_name}.

Voici les {count} derniers emails reçus dans sa boîte de réception :

{emails_json}

Identifie UNIQUEMENT les emails qui nécessitent une réponse de la part de {user_name}.
Ignore : newsletters, notifications automatiques, confirmations, emails en copie sans action requise.

Pour chaque email nécessitant une réponse, génère un brouillon de réponse professionnel en français.

Réponds UNIQUEMENT avec ce JSON :
[
  {{
    "message_id": "id de l'email",
    "subject": "sujet de l'email original",
    "from": "expéditeur",
    "reason": "pourquoi cet email nécessite une réponse",
    "draft_body": "corps du brouillon de réponse (texte brut, professionnel, concis)"
  }}
]

Si aucun email ne nécessite de réponse, retourne [].
"""


def run_drafts(user_email: str | None = None, user_name: str | None = None, dry_run: bool = False) -> list[dict]:
    """
    Main entry point.
    - Reads inbox
    - Asks Claude which emails need a response
    - Creates drafts in Outlook (isDraft=True, never sent)
    - Returns list of created drafts info
    """
    user_email = user_email or os.environ["SENDER_EMAIL"]
    user_name = user_name or os.environ.get("USER_DISPLAY_NAME", user_email.split("@")[0])

    graph = GraphClient()
    messages = graph.list_inbox_messages(user_email, top=30)

    if not messages:
        print("  Aucun email dans la boîte de réception.")
        return []

    print(f"  {len(messages)} email(s) récupéré(s)")

    # Prepare a compact summary for Claude
    emails_summary = [
        {
            "id": m["id"],
            "subject": m.get("subject", "(sans objet)"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address", ""),
            "from_name": m.get("from", {}).get("emailAddress", {}).get("name", ""),
            "received": m.get("receivedDateTime", ""),
            "is_read": m.get("isRead", True),
            "importance": m.get("importance", "normal"),
            "preview": m.get("bodyPreview", "")[:300],
        }
        for m in messages
    ]

    prompt = TRIAGE_PROMPT.format(
        user_name=user_name,
        count=len(emails_summary),
        emails_json=json.dumps(emails_summary, ensure_ascii=False, indent=2),
    )

    client = anthropic.Anthropic()
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)

    drafts_to_create = json.loads(raw)

    if not drafts_to_create:
        print("  Aucun email ne nécessite de réponse.")
        return []

    print(f"  {len(drafts_to_create)} brouillon(s) à créer")

    created = []
    for item in drafts_to_create:
        msg_id = item["message_id"]
        subject = item["subject"]
        draft_body = item["draft_body"]
        sender_addr = item["from"]

        if dry_run:
            print(f"  [DRY RUN] Brouillon pour : {subject}")
            created.append({**item, "draft_id": "dry-run"})
            continue

        try:
            draft_id = graph.create_draft(
                user_email=user_email,
                to_address=sender_addr,
                subject=f"Re: {subject}" if not subject.lower().startswith("re:") else subject,
                body=draft_body,
                reply_to_id=msg_id,
            )
            print(f"  ✓ Brouillon créé : {subject} → {sender_addr}")
            created.append({**item, "draft_id": draft_id})
        except Exception as e:
            print(f"  ✗ Erreur brouillon ({subject}) : {e}")

    return created

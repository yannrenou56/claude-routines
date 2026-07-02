"""Generate meeting summary and personalized action plans using Claude."""

import json
import anthropic
from .transcript_parser import Transcript


def _client() -> anthropic.Anthropic:
    return anthropic.Anthropic()


SUMMARY_PROMPT = """Tu es un assistant expert en gestion de réunions.

Voici le transcript d'une réunion Teams intitulée "{title}".
Participants : {speakers}

TRANSCRIPT :
{transcript}

Génère une réponse JSON avec exactement cette structure :
{{
  "synthese": "Synthèse exécutive en 3-5 phrases. Points clés abordés, décisions prises.",
  "points_cles": ["point 1", "point 2", "point 3"],
  "decisions": ["décision 1", "décision 2"],
  "canal_slack": "nom-du-canal-sans-#",
  "actions_par_participant": {{
    "Nom Participant": ["action 1", "action 2"]
  }}
}}

Pour canal_slack : déduis un nom de canal court en kebab-case depuis le titre de la réunion
(ex: "Réunion Projet Alpha" → "projet-alpha"). Utilise uniquement des lettres minuscules, chiffres et tirets.
Pour actions_par_participant : liste UNIQUEMENT les participants ayant des actions concrètes à faire.
Réponds UNIQUEMENT avec le JSON, sans texte avant ou après."""


def generate_summary(transcript: Transcript) -> dict:
    """Call Claude to produce structured meeting output."""
    prompt = SUMMARY_PROMPT.format(
        title=transcript.meeting_title,
        speakers=", ".join(transcript.speakers),
        transcript=transcript.as_text()[:30000],  # safety trim
    )

    client = _client()
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = message.content[0].text.strip()
    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)

    return json.loads(raw)


def build_email_html(participant: str, summary: dict, meeting_title: str) -> str:
    """Build a personalized HTML email body."""
    actions = summary.get("actions_par_participant", {}).get(participant, [])
    actions_html = ""
    if actions:
        items = "".join(f"<li>{a}</li>" for a in actions)
        actions_html = f"""
        <h2 style="color:#0078D4;">Vos actions</h2>
        <ul>{items}</ul>"""

    points = "".join(f"<li>{p}</li>" for p in summary.get("points_cles", []))
    decisions = "".join(f"<li>{d}</li>" for d in summary.get("decisions", []))

    return f"""
    <html><body style="font-family:Segoe UI,Arial,sans-serif;color:#333;max-width:700px;margin:auto">
      <h1 style="color:#0078D4;border-bottom:2px solid #0078D4;padding-bottom:8px">
        Compte-rendu — {meeting_title}
      </h1>
      <h2>Synthèse</h2>
      <p>{summary.get("synthese", "")}</p>
      <h2>Points clés</h2><ul>{points}</ul>
      <h2>Décisions prises</h2><ul>{decisions}</ul>
      {actions_html}
      <hr style="margin-top:32px"/>
      <p style="font-size:12px;color:#888">Généré automatiquement depuis le transcript Teams</p>
    </body></html>"""


# avoid importing re at module level for a lighter import chain
import re

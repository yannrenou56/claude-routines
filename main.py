#!/usr/bin/env python3
"""
Teams Transcript Processor
Surveille un dossier SharePoint, génère synthèse + plans d'action,
envoie par email à chaque participant.

Usage:
    python main.py             # traite tous les nouveaux transcripts
    python main.py <fichier>   # traite un fichier local directement
"""

import json
import os
import sys
import tempfile
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

from src.transcript_parser import parse_file, Transcript
from src.summarizer import generate_summary, build_email_html
from src.graph_client import GraphClient


def load_processed(path: str) -> set[str]:
    try:
        return set(json.loads(Path(path).read_text()))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def save_processed(path: str, ids: set[str]) -> None:
    Path(path).write_text(json.dumps(sorted(ids), indent=2))


def process_transcript(transcript: Transcript, graph: GraphClient) -> None:
    print(f"\n{'='*60}")
    print(f"Traitement : {transcript.meeting_title}")
    print(f"Participants : {', '.join(transcript.speakers)}")

    print("  → Génération de la synthèse avec Claude...")
    summary = generate_summary(transcript)

    # Emails
    for speaker in transcript.speakers:
        email = graph.resolve_participant_email(speaker)
        if not email:
            print(f"  ⚠ Email introuvable pour {speaker}, ignoré")
            continue
        html = build_email_html(speaker, summary, transcript.meeting_title)
        graph.send_email(
            to_address=email,
            subject=f"Compte-rendu : {transcript.meeting_title}",
            html_body=html,
        )
        print(f"  → Email envoyé à {speaker} <{email}>")

    print(f"  ✓ Terminé")


def run_local_file(file_path: str) -> None:
    graph = GraphClient()
    transcript = parse_file(file_path)
    process_transcript(transcript, graph)


def run_sharepoint_watch() -> None:
    processed_path = os.environ.get("PROCESSED_TRANSCRIPTS_FILE", ".processed_transcripts.json")
    processed_ids = load_processed(processed_path)

    graph = GraphClient()

    print("Récupération des fichiers SharePoint...")
    files = graph.list_transcript_files()
    print(f"{len(files)} fichier(s) trouvé(s) dans le dossier SharePoint")

    new_files = [f for f in files if f["id"] not in processed_ids]
    print(f"{len(new_files)} nouveau(x) transcript(s) à traiter")

    for file_info in new_files:
        name = file_info["name"]
        suffix = Path(name).suffix.lower()
        download_url = file_info.get("@microsoft.graph.downloadUrl", "")

        if not download_url:
            print(f"  ⚠ Pas d'URL de téléchargement pour {name}")
            continue

        try:
            tmp_path = graph.download_file(download_url, suffix)
            transcript = parse_file(tmp_path)
            transcript.meeting_title = transcript.meeting_title or Path(name).stem
            process_transcript(transcript, graph)
            processed_ids.add(file_info["id"])
            save_processed(processed_path, processed_ids)
        except Exception as e:
            print(f"  ✗ Erreur sur {name} : {e}")
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_local_file(sys.argv[1])
    else:
        run_sharepoint_watch()

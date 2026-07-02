"""Post meeting summaries to the right Slack channel."""

import os
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError


class SlackNotifier:
    def __init__(self):
        self._client = WebClient(token=os.environ["SLACK_BOT_TOKEN"])
        self._channel_cache: dict[str, str | None] = {}

    def _find_channel_id(self, channel_name: str) -> str | None:
        """Resolve a channel name to its Slack ID."""
        if channel_name in self._channel_cache:
            return self._channel_cache[channel_name]

        name = channel_name.lstrip("#")
        cursor = None
        while True:
            kwargs = {"exclude_archived": True, "limit": 200}
            if cursor:
                kwargs["cursor"] = cursor
            resp = self._client.conversations_list(**kwargs)
            for ch in resp["channels"]:
                if ch["name"] == name:
                    self._channel_cache[channel_name] = ch["id"]
                    return ch["id"]
            cursor = resp.get("response_metadata", {}).get("next_cursor")
            if not cursor:
                break

        self._channel_cache[channel_name] = None
        return None

    def _get_or_create_channel(self, channel_name: str) -> str:
        """Find or create a channel, returning its ID."""
        channel_id = self._find_channel_id(channel_name)
        if channel_id:
            return channel_id

        # Try to create the channel
        try:
            resp = self._client.conversations_create(name=channel_name.lstrip("#"), is_private=False)
            return resp["channel"]["id"]
        except SlackApiError as e:
            if "name_taken" in str(e):
                channel_id = self._find_channel_id(channel_name)
                if channel_id:
                    return channel_id
            raise

    def post_summary(self, channel_name: str, meeting_title: str, summary: dict) -> str:
        """Post the meeting summary to a Slack channel. Returns the message ts."""
        channel_id = self._get_or_create_channel(channel_name)

        points = "\n".join(f"• {p}" for p in summary.get("points_cles", []))
        decisions = "\n".join(f"✅ {d}" for d in summary.get("decisions", []))

        blocks = [
            {
                "type": "header",
                "text": {"type": "plain_text", "text": f"Compte-rendu : {meeting_title}"},
            },
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": f"*Synthèse*\n{summary.get('synthese', '')}"},
            },
        ]
        if points:
            blocks.append({
                "type": "section",
                "text": {"type": "mrkdwn", "text": f"*Points clés*\n{points}"},
            })
        if decisions:
            blocks.append({
                "type": "section",
                "text": {"type": "mrkdwn", "text": f"*Décisions*\n{decisions}"},
            })
        blocks.append({"type": "divider"})

        resp = self._client.chat_postMessage(channel=channel_id, blocks=blocks, text=f"CR: {meeting_title}")
        return resp["ts"]

    def post_action_plan_thread(self, channel_name: str, thread_ts: str, participant: str, actions: list[str]) -> None:
        """Post a participant's action plan as a thread reply."""
        if not actions:
            return
        channel_id = self._find_channel_id(channel_name)
        if not channel_id:
            return
        action_text = "\n".join(f"• {a}" for a in actions)
        self._client.chat_postMessage(
            channel=channel_id,
            thread_ts=thread_ts,
            text=f"*Actions — {participant}*\n{action_text}",
        )

    def default_channel(self) -> str:
        return os.environ.get("SLACK_DEFAULT_CHANNEL", "#general")

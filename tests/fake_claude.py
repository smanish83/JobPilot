"""A stand-in for the Claude Code CLI (`claude`), used by test_auth.py.

  claude auth status        -> JSON from FAKE_CLAUDE_STATUS (exit 0 when signed in, 1 when not)
  claude -p ... < prompt    -> {"type": "result", "result": <reply>, ...}; the reply is chosen by the start of the
                               --system-prompt-file text, from the JSON map in FAKE_CLAUDE_REPLIES
Every call is appended to FAKE_CLAUDE_LOG as one JSON line (args and working folder).
"""
import json
import os
import sys

args = sys.argv[1:]
with open(os.environ["FAKE_CLAUDE_LOG"], "a", encoding="utf-8") as log:
    log.write(json.dumps({"args": args, "cwd": os.getcwd(), "files": sorted(os.listdir("."))}) + "\n")

if args[:2] == ["auth", "status"]:
    status = json.loads(os.environ.get("FAKE_CLAUDE_STATUS", '{"loggedIn": false, "authMethod": "none"}'))
    print(json.dumps(status))
    sys.exit(0 if status.get("loggedIn") else 1)

if os.environ.get("FAKE_CLAUDE_OLD") and "--tools" in args:
    sys.stderr.write("error: unknown option '--tools'\n")
    sys.exit(1)

prompt = sys.stdin.read()
system = open(args[args.index("--system-prompt-file") + 1], encoding="utf-8").read()
if os.environ.get("FAKE_CLAUDE_FAIL"):
    print(json.dumps({"type": "result", "is_error": True, "result": os.environ["FAKE_CLAUDE_FAIL"]}))
    sys.exit(1)
replies = json.loads(open(os.environ["FAKE_CLAUDE_REPLIES"], encoding="utf-8").read())
reply = next((r for prefix, r in replies.items() if system.startswith(prefix)), None)
if reply is None:
    print(json.dumps({"type": "result", "is_error": True, "result": "unexpected system prompt"}))
    sys.exit(1)
print(json.dumps({"type": "result", "subtype": "success", "is_error": False, "stop_reason": "end_turn",
                  "result": json.dumps(reply), "prompt_chars": len(prompt)}))

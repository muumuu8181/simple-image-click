# -*- coding: utf-8 -*-
"""
Claude's query to other AI via the automation system
Usage: python claude_query.py "your prompt here" [flow_name]
"""
import requests
import json
import time
import sys
from pathlib import Path

BASE_URL = "http://localhost:8000"

def add_text(text):
    """Add a text and get its ID"""
    res = requests.post(f"{BASE_URL}/api/texts", json={"text": text})
    return res.json()

def get_flow(flow_name):
    """Get flow by name"""
    res = requests.get(f"{BASE_URL}/api/flows")
    flows = res.json().get("flows", {})
    return flows.get(flow_name)

def execute_flow(actions, text_id, flow_name, group_name):
    """Execute flow with modified text_id"""
    # Replace text_id in paste and save_to_file actions
    modified_actions = []
    for a in actions:
        if a.get("type") == "paste":
            modified_actions.append({**a, "text_id": text_id})
        elif a.get("type") == "save_to_file":
            modified_actions.append({**a, "text_id": text_id, "flow_name": flow_name, "group_name": group_name})
        else:
            modified_actions.append(a)

    res = requests.post(f"{BASE_URL}/api/execute", json={
        "actions": modified_actions,
        "interval": 2,
        "confidence": 0.95,
        "min_confidence": 0.7,
        "wait_timeout": 1800,
        "cursor_speed": 0.5,
        "start_delay": 3  # Give me time to switch windows
    })
    return res.json()

def poll_status():
    """Poll execution status until complete"""
    while True:
        res = requests.get(f"{BASE_URL}/api/execute/status")
        status = res.json()
        print(f"Progress: {status.get('current_step')}/{status.get('total_steps')}")
        if not status.get("is_running"):
            return status
        time.sleep(2)

if __name__ == "__main__":
    # Parse arguments
    if len(sys.argv) < 2:
        print("Usage: python claude_query.py \"your prompt\" [flow_name]")
        print("  flow_name defaults to: 通常プロンプト-Liner")
        exit(1)

    my_question = sys.argv[1]
    flow_name = sys.argv[2] if len(sys.argv) > 2 else "通常プロンプト-Liner"

    # Grok系フローの場合、日本語で回答するよう指示を追加
    if "Grok" in flow_name or "grok" in flow_name:
        my_question = my_question + "

※日本語で回答してください。"
        print("[Grok] 日本語回答指示を追加")

    print(f"Prompt: {my_question[:50]}...")
    result = add_text(my_question)
    text_id = result.get("id")
    print(f"Text ID: {text_id}")

    # Get flow
    flow = get_flow(flow_name)
    if not flow:
        print(f"Flow not found: {flow_name}")
        exit(1)

    print(f"Flow: {flow_name} ({len(flow['actions'])} actions)")

    # Execute
    print("Executing... (3 second delay)")
    print(">>> SWITCH TO BROWSER NOW <<<")
    exec_result = execute_flow(flow["actions"], text_id, flow_name, flow.get("group", "ai-normal"))
    print(f"Started: {exec_result}")

    # Poll for completion
    final_status = poll_status()
    print(f"Completed: {final_status.get('current_step')}/{final_status.get('total_steps')} steps")

    # Show results
    for r in final_status.get("results", []):
        status_icon = "OK" if r.get("status") == "success" else "NG"
        print(f"{status_icon} {r.get('message')}")

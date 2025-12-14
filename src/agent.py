"""LLM Agent for task analysis and decomposition using Claude."""

import json
from typing import Optional
from dataclasses import dataclass
from anthropic import Anthropic

from .lark_client import LarkClient


@dataclass
class AgentAction:
    """Represents an action the agent wants to take."""
    action_type: str  # "create_task", "create_subtasks", "complete_task", "list_tasks", "respond"
    parameters: dict
    reasoning: str


SYSTEM_PROMPT = """あなたはLarkタスク管理のAIアシスタントです。
ユーザーの音声入力を理解し、適切なタスク管理操作を行います。

## あなたの能力:
1. タスクの作成 - 新しいタスクを作成
2. サブタスク分解 - 複雑なタスクを具体的なサブタスクに分解
3. タスク完了 - タスクを完了としてマーク
4. タスク一覧 - 現在のタスクを表示
5. 一般的な応答 - タスク管理に関する質問に回答

## ツール:
以下のツールを使用できます。必ずJSON形式で応答してください。

### create_task
新しいタスクを作成します。
```json
{
  "action": "create_task",
  "parameters": {
    "summary": "タスクのタイトル",
    "description": "タスクの詳細説明（オプション）",
    "due": "期限（ISO 8601形式、オプション）"
  },
  "reasoning": "なぜこのタスクを作成するか"
}
```

### create_subtasks
タスクをサブタスクに分解します。親タスクがない場合は先に親タスクを作成します。
```json
{
  "action": "create_subtasks",
  "parameters": {
    "parent_summary": "親タスクのタイトル（新規作成する場合）",
    "parent_task_guid": "既存の親タスクのGUID（既存タスクに追加する場合）",
    "subtasks": [
      {"summary": "サブタスク1", "description": "詳細"},
      {"summary": "サブタスク2", "description": "詳細"}
    ]
  },
  "reasoning": "タスク分解の理由"
}
```

### complete_task
タスクを完了としてマークします。
```json
{
  "action": "complete_task",
  "parameters": {
    "task_identifier": "タスク名またはGUID"
  },
  "reasoning": "完了理由"
}
```

### list_tasks
タスク一覧を取得します。
```json
{
  "action": "list_tasks",
  "parameters": {},
  "reasoning": "一覧取得の理由"
}
```

### respond
ユーザーに応答のみ行い、操作は行いません。
```json
{
  "action": "respond",
  "parameters": {
    "message": "ユーザーへのメッセージ"
  },
  "reasoning": "応答のみの理由"
}
```

## 重要なルール:
1. 常に1つのJSON応答のみを返す
2. タスク分解時は具体的で実行可能なサブタスクを作成
3. ユーザーの意図が不明な場合は確認の応答を返す
4. 日本語で応答する
"""


class TaskAgent:
    """Agent that analyzes user input and manages Lark tasks."""

    def __init__(self, anthropic_api_key: str, lark_client: LarkClient):
        self.client = Anthropic(api_key=anthropic_api_key)
        self.lark = lark_client
        self.conversation_history: list[dict] = []
        self.default_tasklist_guid: Optional[str] = None

    def set_default_tasklist(self, tasklist_guid: str):
        """Set the default task list for new tasks."""
        self.default_tasklist_guid = tasklist_guid

    def _parse_action(self, response_text: str) -> AgentAction:
        """Parse the agent's response into an action."""
        # Try to extract JSON from the response
        try:
            # Look for JSON in the response
            start = response_text.find('{')
            end = response_text.rfind('}') + 1
            if start >= 0 and end > start:
                json_str = response_text[start:end]
                data = json.loads(json_str)
                return AgentAction(
                    action_type=data.get("action", "respond"),
                    parameters=data.get("parameters", {}),
                    reasoning=data.get("reasoning", "")
                )
        except json.JSONDecodeError:
            pass

        # If parsing fails, treat as a simple response
        return AgentAction(
            action_type="respond",
            parameters={"message": response_text},
            reasoning="JSON解析に失敗したため、テキスト応答として処理"
        )

    def _execute_action(self, action: AgentAction) -> str:
        """Execute the parsed action and return result message."""
        try:
            if action.action_type == "create_task":
                params = action.parameters
                task = self.lark.create_task(
                    summary=params.get("summary", "新しいタスク"),
                    description=params.get("description", ""),
                    due=params.get("due"),
                    tasklist_guid=self.default_tasklist_guid
                )
                return f"✅ タスクを作成しました: {task.get('summary', params.get('summary'))}"

            elif action.action_type == "create_subtasks":
                params = action.parameters
                parent_guid = params.get("parent_task_guid")

                # Create parent task if not specified
                if not parent_guid and params.get("parent_summary"):
                    parent_task = self.lark.create_task(
                        summary=params["parent_summary"],
                        tasklist_guid=self.default_tasklist_guid
                    )
                    parent_guid = parent_task.get("guid")

                if not parent_guid:
                    return "❌ 親タスクが指定されていません"

                # Create subtasks
                subtasks = params.get("subtasks", [])
                created = []
                for subtask in subtasks:
                    result = self.lark.create_subtask(
                        parent_task_guid=parent_guid,
                        summary=subtask.get("summary", ""),
                        description=subtask.get("description", "")
                    )
                    created.append(result.get("summary", subtask.get("summary")))

                return f"✅ {len(created)}個のサブタスクを作成しました:\n" + "\n".join(f"  • {s}" for s in created)

            elif action.action_type == "complete_task":
                identifier = action.parameters.get("task_identifier", "")

                # Try to find task by name
                tasks = self.lark.get_tasks(tasklist_guid=self.default_tasklist_guid)
                matched_task = None
                for task in tasks:
                    if identifier.lower() in task.get("summary", "").lower():
                        matched_task = task
                        break

                if matched_task:
                    self.lark.complete_task(matched_task["guid"])
                    return f"✅ タスクを完了しました: {matched_task.get('summary')}"
                else:
                    return f"❌ タスクが見つかりません: {identifier}"

            elif action.action_type == "list_tasks":
                tasks = self.lark.get_tasks(tasklist_guid=self.default_tasklist_guid)
                if not tasks:
                    return "📋 タスクはありません"

                lines = ["📋 タスク一覧:"]
                for task in tasks[:10]:  # Limit to 10 tasks
                    status = "✓" if task.get("completed_at") else "○"
                    lines.append(f"  {status} {task.get('summary', '無題')}")
                return "\n".join(lines)

            elif action.action_type == "respond":
                return action.parameters.get("message", "")

            else:
                return f"⚠️ 未知のアクション: {action.action_type}"

        except Exception as e:
            return f"❌ エラーが発生しました: {str(e)}"

    def process_input(self, user_input: str) -> str:
        """Process user input and return the result.

        Args:
            user_input: User's voice/text input

        Returns:
            Result message to show to user
        """
        # Add user message to history
        self.conversation_history.append({
            "role": "user",
            "content": user_input
        })

        # Call Claude API
        response = self.client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=self.conversation_history
        )

        assistant_message = response.content[0].text

        # Add assistant response to history
        self.conversation_history.append({
            "role": "assistant",
            "content": assistant_message
        })

        # Parse and execute action
        action = self._parse_action(assistant_message)

        # Execute the action
        result = self._execute_action(action)

        return result

    def analyze_and_decompose_task(self, task_description: str) -> str:
        """Analyze a task and decompose it into subtasks.

        This is a specialized method for task decomposition.
        """
        prompt = f"""以下のタスクを分析し、具体的で実行可能なサブタスクに分解してください。

タスク: {task_description}

サブタスクは以下の基準で作成してください：
1. 各サブタスクは1時間以内で完了できる具体的なアクション
2. 順序が重要な場合は順番に並べる
3. 依存関係がある場合は明記する
4. 完了条件が明確であること

create_subtasks アクションを使用してJSON形式で応答してください。"""

        return self.process_input(prompt)

    def clear_history(self):
        """Clear conversation history."""
        self.conversation_history = []

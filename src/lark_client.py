"""Lark API client for task management."""

import httpx
from typing import Optional
from dataclasses import dataclass
import time


@dataclass
class LarkTask:
    """Represents a Lark task."""
    id: str
    summary: str
    description: str = ""
    due: Optional[str] = None
    completed: bool = False


class LarkClient:
    """Client for interacting with Lark Open Platform API."""

    BASE_URL = "https://open.larksuite.com/open-apis"

    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self._access_token: Optional[str] = None
        self._token_expires_at: float = 0
        self._client = httpx.Client(timeout=30.0)

    def _get_access_token(self) -> str:
        """Get or refresh the tenant access token."""
        if self._access_token and time.time() < self._token_expires_at:
            return self._access_token

        response = self._client.post(
            f"{self.BASE_URL}/auth/v3/tenant_access_token/internal",
            json={
                "app_id": self.app_id,
                "app_secret": self.app_secret,
            }
        )
        response.raise_for_status()
        data = response.json()

        if data.get("code") != 0:
            raise Exception(f"Failed to get access token: {data.get('msg')}")

        self._access_token = data["tenant_access_token"]
        self._token_expires_at = time.time() + data.get("expire", 7200) - 300
        return self._access_token

    def _request(self, method: str, endpoint: str, **kwargs) -> dict:
        """Make an authenticated request to Lark API."""
        token = self._get_access_token()
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"

        response = self._client.request(
            method,
            f"{self.BASE_URL}{endpoint}",
            headers=headers,
            **kwargs
        )
        response.raise_for_status()
        return response.json()

    def get_tasklists(self, page_size: int = 50) -> list[dict]:
        """Get all task lists."""
        result = self._request(
            "GET",
            "/task/v2/tasklists",
            params={"page_size": page_size}
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to get tasklists: {result.get('msg')}")
        return result.get("data", {}).get("items", [])

    def create_task(
        self,
        summary: str,
        description: str = "",
        due: Optional[str] = None,
        tasklist_guid: Optional[str] = None,
    ) -> dict:
        """Create a new task.

        Args:
            summary: Task title
            description: Task description
            due: Due date in RFC3339 format (e.g., "2024-12-31T23:59:59+09:00")
            tasklist_guid: Optional task list GUID to add the task to
        """
        task_data = {
            "summary": summary,
        }
        if description:
            task_data["description"] = description
        if due:
            task_data["due"] = {"timestamp": due}

        result = self._request(
            "POST",
            "/task/v2/tasks",
            json=task_data
        )

        if result.get("code") != 0:
            raise Exception(f"Failed to create task: {result.get('msg')}")

        task = result.get("data", {}).get("task", {})

        # Add to task list if specified
        if tasklist_guid and task.get("guid"):
            self.add_task_to_tasklist(task["guid"], tasklist_guid)

        return task

    def add_task_to_tasklist(self, task_guid: str, tasklist_guid: str) -> dict:
        """Add a task to a task list."""
        result = self._request(
            "POST",
            f"/task/v2/tasks/{task_guid}/add_tasklist",
            json={"tasklist_guid": tasklist_guid}
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to add task to list: {result.get('msg')}")
        return result.get("data", {})

    def create_subtask(
        self,
        parent_task_guid: str,
        summary: str,
        description: str = "",
    ) -> dict:
        """Create a subtask under a parent task."""
        result = self._request(
            "POST",
            f"/task/v2/tasks/{parent_task_guid}/subtasks",
            json={
                "summary": summary,
                "description": description,
            }
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to create subtask: {result.get('msg')}")
        return result.get("data", {}).get("subtask", {})

    def get_tasks(
        self,
        tasklist_guid: Optional[str] = None,
        page_size: int = 50
    ) -> list[dict]:
        """Get tasks, optionally filtered by task list."""
        params = {"page_size": page_size}

        if tasklist_guid:
            # Get tasks from specific task list
            result = self._request(
                "GET",
                f"/task/v2/tasklists/{tasklist_guid}/tasks",
                params=params
            )
        else:
            # Get all tasks
            result = self._request(
                "GET",
                "/task/v2/tasks",
                params=params
            )

        if result.get("code") != 0:
            raise Exception(f"Failed to get tasks: {result.get('msg')}")
        return result.get("data", {}).get("items", [])

    def get_task(self, task_guid: str) -> dict:
        """Get a specific task by GUID."""
        result = self._request("GET", f"/task/v2/tasks/{task_guid}")
        if result.get("code") != 0:
            raise Exception(f"Failed to get task: {result.get('msg')}")
        return result.get("data", {}).get("task", {})

    def complete_task(self, task_guid: str) -> dict:
        """Mark a task as completed."""
        result = self._request(
            "POST",
            f"/task/v2/tasks/{task_guid}/complete"
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to complete task: {result.get('msg')}")
        return result.get("data", {})

    def uncomplete_task(self, task_guid: str) -> dict:
        """Mark a task as not completed."""
        result = self._request(
            "POST",
            f"/task/v2/tasks/{task_guid}/uncomplete"
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to uncomplete task: {result.get('msg')}")
        return result.get("data", {})

    def update_task(
        self,
        task_guid: str,
        summary: Optional[str] = None,
        description: Optional[str] = None,
        due: Optional[str] = None,
    ) -> dict:
        """Update a task."""
        update_fields = []
        task_data = {}

        if summary is not None:
            task_data["summary"] = summary
            update_fields.append("summary")
        if description is not None:
            task_data["description"] = description
            update_fields.append("description")
        if due is not None:
            task_data["due"] = {"timestamp": due}
            update_fields.append("due")

        if not update_fields:
            raise ValueError("No fields to update")

        result = self._request(
            "PATCH",
            f"/task/v2/tasks/{task_guid}",
            json=task_data,
            params={"update_fields": ",".join(update_fields)}
        )
        if result.get("code") != 0:
            raise Exception(f"Failed to update task: {result.get('msg')}")
        return result.get("data", {}).get("task", {})

    def delete_task(self, task_guid: str) -> bool:
        """Delete a task."""
        result = self._request("DELETE", f"/task/v2/tasks/{task_guid}")
        if result.get("code") != 0:
            raise Exception(f"Failed to delete task: {result.get('msg')}")
        return True

    def close(self):
        """Close the HTTP client."""
        self._client.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

"""接口层测试。"""

from __future__ import annotations

from pathlib import Path


def test_health(api_client) -> None:
    response = api_client.get("/api/v1/health")
    assert response.status_code == 200

    body = response.json()
    assert body["status"] == "ok"
    assert "engine" in body


def test_root(api_client) -> None:
    response = api_client.get("/")
    assert response.status_code == 200
    assert response.json()["api"] == "/api/v1"


def test_engine_capabilities(api_client) -> None:
    response = api_client.get("/api/v1/engines")
    assert response.status_code == 200

    body = response.json()
    assert "available" in body
    assert "refresh_fields" in body


def test_convert_sync_returns_docx(api_client, sample_docx: Path) -> None:
    with sample_docx.open("rb") as fp:
        response = api_client.post(
            "/api/v1/documents/convert-sync",
            files={"file": (sample_docx.name, fp, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            data={"toc_title": "目      录"},
        )

    assert response.status_code == 200, response.text
    assert len(response.content) > 1000
    assert "wordprocessingml" in response.headers["content-type"]


def test_convert_sync_rejects_wrong_extension(api_client) -> None:
    response = api_client.post(
        "/api/v1/documents/convert-sync",
        files={"file": ("bad.txt", b"hello", "text/plain")},
    )
    assert response.status_code == 400
    assert response.json()["code"] == "validation_error"


def test_convert_async_creates_task(api_client, sample_docx: Path) -> None:
    with sample_docx.open("rb") as fp:
        response = api_client.post(
            "/api/v1/documents/convert",
            files={"file": (sample_docx.name, fp, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        )

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["task_id"]
    assert body["events_url"].endswith("/events")

    # 任务应能被查询到
    detail = api_client.get(f"/api/v1/tasks/{body['task_id']}")
    assert detail.status_code == 200
    assert detail.json()["task_id"] == body["task_id"]


def test_unknown_task_returns_404(api_client) -> None:
    response = api_client.get("/api/v1/tasks/not-exists")
    assert response.status_code == 404
    assert response.json()["code"] == "task_not_found"


def test_template_register_and_list(api_client, template_docx: Path) -> None:
    with template_docx.open("rb") as fp:
        created = api_client.post(
            "/api/v1/templates",
            files={"file": (template_docx.name, fp, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            data={"name": "测试模板", "kind": "content", "is_default": "true"},
        )

    assert created.status_code == 200, created.text
    template_id = created.json()["template_id"]

    listed = api_client.get("/api/v1/templates")
    assert listed.status_code == 200
    assert any(item["template_id"] == template_id for item in listed.json())

    deleted = api_client.delete(f"/api/v1/templates/{template_id}")
    assert deleted.status_code == 200


def test_history_lists_records(api_client, sample_docx: Path) -> None:
    with sample_docx.open("rb") as fp:
        api_client.post(
            "/api/v1/documents/convert-sync",
            files={"file": (sample_docx.name, fp, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        )

    response = api_client.get("/api/v1/history")
    assert response.status_code == 200
    assert "total" in response.json()

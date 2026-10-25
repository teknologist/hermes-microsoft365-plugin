def test_every_executable_operation_has_handler_metadata_and_sdk_case():
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE
    from microsoft365 import sdk_contract

    for key, definition in OPERATION_REGISTRY.items():
        if definition.executable:
            assert key in HANDLER_TABLE, key
            assert definition.endpoints, key
            assert definition.contract_cases, key
            for call_name, case in definition.contract_cases:
                if not call_name:
                    continue
                assert hasattr(sdk_contract, call_name), (key, call_name)
                assert case in {"outlook_attachments", "outlook_attachment", "outlook_messages", "outlook_read", "calendar_events", "calendar_read", "todo_lists", "todo_tasks", "todo_read", "planner_plans", "planner_buckets", "planner_tasks", "planner_read", "drive_search", "drive_item_read", "download_files", "sites_search", "teams_joined", "teams_channels"}, (key, case)


def test_registry_executable_set_is_exact_and_writes_are_withheld():
    from microsoft365.contract import EXECUTABLE_OPERATIONS, WRITE_OPERATIONS
    from microsoft365.registration import active_actions
    from microsoft365.contract import Settings

    expected = {
        "outlook.search", "outlook.read", "outlook.list_attachments", "outlook.download_attachment", "calendar.search",
        "sharepoint.search", "sharepoint.read", "sharepoint.download_files",
        "onedrive.search", "onedrive.read", "onedrive.download_files",
        "teams.list_teams", "teams.list_channels",
        "todo.list_task_lists", "todo.search", "todo.read",
        "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read",
    }
    assert set(EXECUTABLE_OPERATIONS) == expected
    settings = Settings.from_mapping({"capabilities": {service: True for service in ("todo", "planner")}})
    assert not set(active_actions(settings, "todo")) & WRITE_OPERATIONS
    assert not set(active_actions(settings, "planner")) & WRITE_OPERATIONS


def test_delegated_mode_is_not_implemented_for_new_reads():
    from microsoft365.contract import operation_status
    for service, operation in (("todo", "read"), ("planner", "read")):
        status = operation_status("delegated", service, operation)
        assert status.executable is False
        assert status.auth_status == "not_implemented"

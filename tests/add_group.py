

def test_add_group(app, xls_groups):
    group = xls_groups
    main_window = app.get_main_window()
    old_groups = app.group.get_groups(main_window)
    app.group.add(main_window, group)
    new_groups = app.group.get_groups(main_window)
    old_groups.append(group)
    assert sorted(old_groups) == sorted(new_groups)

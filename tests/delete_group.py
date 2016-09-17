import random


def test_add_group(app):
    main_window = app.get_main_window()
    if len(app.group.get_groups(main_window)) == 0:
        app.group.add(main_window, 'new')
    old_groups = app.group.get_groups(main_window)
    group = random.choice(old_groups)
    app.group.delete(main_window, group)
    new_groups = app.group.get_groups(main_window)
    old_groups.remove(group)
    assert sorted(old_groups) == sorted(new_groups)

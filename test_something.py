import clr
import os.path

project = os.path.dirname(os.path.abspath(__file__))
import sys
sys.path.append(os.path.join(project, 'TestStack.White.9.2.0.11\\lib\\net40\\'))
sys.path.append(os.path.join(project, 'Castle.Core.3.1.0\\lib\\net40-client\\'))
clr.AddReferenceByName('TestStack.White')

from TestStack.White import Application
from TestStack.White.UIItems.Finders import *
from TestStack.White.InputDevices import KeyBoard
from TestStack.White.WindowsAPI import KeyBoardInput

clr.AddReferenceByName('UIAutomationTypes, Version=3.0.0.0, Culture=Neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *


def get_groups(main_window):
    modal = open_group_editor(main_window)
    tree = modal.Get(SearchCriteria.ByAutomationId('uxAddressTreeView'))
    root = tree.Nodes[0]
    groups = [node.Text for node in root.Nodes]
    close_group_editor(modal)
    return groups


def add(main_window, group_name):
    modal = open_group_editor(main_window)
    modal.Get(SearchCriteria.ByAutomationId('uxNewAddressButton')).Click()
    modal.Get(SearchCriteria.ByControlType(ControlType.Edit)).Enter(group_name)
    KeyBoard.Instance.PressSpecialKey(KeyBoardInput.SpecialKeys.RETURN)
    close_group_editor(modal)


def close_group_editor(modal):
    modal.Get(SearchCriteria.ByAutomationId('uxCloseAddressButton')).Click()


def open_group_editor(main_window):
    main_window.Get(SearchCriteria.ByAutomationId('groupButton')).Click()
    modal = main_window.MoadlWindow('Group Editor')
    return modal


def test_something():
    application = Application.Launch('C:\\D\\GitHub\\iron\\FreeAddressBookPortable\\AddressBook.exe')
    main_window = application.GetWindow('Free Address Book')
    old_groups = get_groups()
    add(main_window, 'NEW')
    new_groups = get_groups()
    old_groups.append('NEW')
    assert sorted(old_groups) == sorted(new_groups)
    main_window.Get(SearchCriteria.ByAutomationId('uxExitAddressButton')).Click()

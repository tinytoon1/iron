import os
import sys

project = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(project, 'TestStack.White.0.13.3\\lib\\net40'))
sys.path.append(os.path.join(project, 'Castle.Core.3.3.0\\lib\\net40-client'))

import clr

clr.AddReferenceByName('TestStack.White')
from TestStack.White.UIItems.Finders import *
from TestStack.White.InputDevices import Keyboard
from TestStack.White.WindowsAPI import KeyboardInput

clr.AddReferenceByName('UIAutomationTypes, Version=3.0.0.0, Culture=Neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *


class GroupHelper:
    def __init__(self, app):
        self.app = app

    def get_groups(self, main_window):
        modal = self.open_group_editor(main_window)
        tree = modal.Get(SearchCriteria.ByAutomationId('uxAddressTreeView'))
        root = tree.Nodes[0]
        groups = [node.Text for node in root.Nodes]
        self.close_group_editor(modal)
        return groups

    def add(self, main_window, group_name):
        modal = self.open_group_editor(main_window)
        modal.Get(SearchCriteria.ByAutomationId('uxNewAddressButton')).Click()
        modal.Get(SearchCriteria.ByControlType(ControlType.Edit)).Enter(group_name)
        Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN)
        self.close_group_editor(modal)

    def delete(self, main_window, group_name):
        modal = self.open_group_editor(main_window)
        modal.Get(SearchCriteria.ByAutomationId('uxNewAddressButton')).Click()
        modal.Get(SearchCriteria.ByName(group_name)).Click()
        modal.Get(SearchCriteria.ByAutomationId('uxDeleteAddressButton')).Click()
        confirm = modal.ModalWindow('Delete Group')
        confirm.Get(SearchCriteria.ByAutomationId('uxOKAddressButton')).Click()
        self.close_group_editor(modal)

    def open_group_editor(self, main_window):
        main_window.Get(SearchCriteria.ByAutomationId('groupButton')).Click()
        modal = main_window.ModalWindow('Group Editor')
        return modal

    def close_group_editor(self, modal):
        modal.Get(SearchCriteria.ByAutomationId('uxCloseAddressButton')).Click()

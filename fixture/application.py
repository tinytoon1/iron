import os
import sys
from fixture.group import GroupHelper

project = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(project, 'TestStack.White.0.13.3\\lib\\net40'))
sys.path.append(os.path.join(project, 'Castle.Core.3.3.0\\lib\\net40-client'))

import clr

clr.AddReferenceByName('TestStack.White')
from TestStack.White import Application


class App:
    def __init__(self, path, title):
        self.group = GroupHelper(self)
        self.path = path
        self.title = title

    def run_application(self):
        self.application = Application.Launch(self.path)

    def get_main_window(self):
        return self.application.GetWindow(self.title)

    def destroy(self):
        self.application.Close()

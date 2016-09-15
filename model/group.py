

class Group:
    def __init__(self, name=None):
        self.name = name

    def __repr__(self):
        return "%s: %s %s %s" % (self.id, self.name, self.header, self.footer)

    def __eq__(self, other):
        return self.name == other.name

    def id_or_max(self):
        return self.nam

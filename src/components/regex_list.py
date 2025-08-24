# Simple class to hold and return regexes


class RegexList:
    def __init__(self, regexes=None):
        if regexes is None:
            # Example: two random regexes
            regexes = [
                r'(?i)\s+if\s+betingelse\s+(.+?)\s*(?=[“”"])[“”"]([^“”"]*)[“”"]\s*else\s*[“”"]([^“”"]*)[“”"]'
                # ,r'(?i)Else til if betingelse\s+(.+?)\s*[“”"]([^“”"]*)[“”"]',
            ]
        self._regexes = regexes

    def get_regexes(self):
        return self._regexes

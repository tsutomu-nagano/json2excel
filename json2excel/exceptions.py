
class MissingRequiredError(Exception):
    def __init__(self, args):
        super().__init__(f"Required '{args}' is missing.")

class InvalidNumericValueError(Exception):
    def __init__(self, cell_name):
        self.cell_name = cell_name
        super().__init__(f"'{cell_name}' is Invalid Numeric Value.")
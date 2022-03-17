class FASTAFormattingError(Exception):
    """
    Raise when content of the provided file
    does not comply with the FASTA formatting standards
    """

    def __init__(self):
        self.message = "Provided file content does not comply with FASTA formatting standards"

    def __str__(self):
        return self.message

    def __repr__(self):
        return self.message

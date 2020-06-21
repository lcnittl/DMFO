from pathlib import Path


class VCSFileData:
    target_ext: str

    def __init__(self, name):
        self.name: Path = name
        self.fileobj: object
        self.is_lfs: bool

    def get_name(self) -> Path:
        """Returns name if it has the target extension, otherwise it returns the name
        appended by the target extension.
        """
        if self.name.suffix != self.target_ext:
            return Path(str(self.name) + self.target_ext)
        return self.name

    def has_ext(self) -> bool:
        """Returns True if it has the target extension, else False"""
        return self.name.suffix == self.target_ext

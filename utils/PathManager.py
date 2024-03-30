import os


class PathManager:
    def __init__(self, root_dir):
        self.root_dir = root_dir

    def get_path(self, *path_parts):
        return os.path.join(self.root_dir, *path_parts)

    class PathWrapper:
        def __init__(self, path_manager, path_parts):
            self.path_manager = path_manager
            self.path_parts = path_parts

        def __getattr__(self, folder_name):
            # new_path_parts = self.path_parts + [folder_name]
            new_path_parts = self.path_parts + [folder_name.replace("_", "-")]  # Replace the underline with a hyphen
            return self.path_manager.PathWrapper(self.path_manager, new_path_parts)

        def __call__(self, *file_parts):
            return self.path_manager.get_path(*self.path_parts, *file_parts)

        def __repr__(self):
            return str(self())

    @property
    def assets(self):
        return self.PathWrapper(self, ["assets"])

    @property
    def input(self):
        return self.PathWrapper(self, ["input"])

    @property
    def output(self):
        return self.PathWrapper(self, ["output"])

    @property
    def transit(self):
        return self.PathWrapper(self, ["transit"])

    @property
    def refer(self):
        return self.PathWrapper(self, ["refer"])

    @property
    def template(self):
        return self.PathWrapper(self, ["template"])

    @property
    def config(self):
        return self.PathWrapper(self, ["config"])

load_path_manager = PathManager(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from typing import List, Optional
from src.core.base_module import BaseDocumentModule

class ModuleRegistry:
    _modules_by_ext = {}
    _modules_by_name = {}

    @classmethod
    def register(cls, module: BaseDocumentModule):
        cls._modules_by_name[module.name] = module
        for ext in module.file_extensions:
            cls._modules_by_ext[ext.lower()] = module

    @classmethod
    def get_module_by_extension(cls, ext: str) -> Optional[BaseDocumentModule]:
        return cls._modules_by_ext.get(ext.lower())

    @classmethod
    def get_module_by_name(cls, name: str) -> Optional[BaseDocumentModule]:
        return cls._modules_by_name.get(name)

    @classmethod
    def get_all_modules(cls) -> List[BaseDocumentModule]:
        return list(cls._modules_by_name.values())

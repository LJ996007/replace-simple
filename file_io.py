"""将完整写入的临时文件提交到最终路径。"""

import os
import tempfile
from contextlib import contextmanager


@contextmanager
def atomic_output_path(output_path):
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(
        prefix=".replace-simple-", suffix=os.path.splitext(output_path)[1],
        dir=os.path.dirname(output_path),
    )
    os.close(descriptor)
    try:
        yield temporary
        os.replace(temporary, output_path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)

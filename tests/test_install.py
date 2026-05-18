"""
测试 install 模块
"""

import unittest
from io import StringIO
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import install


class TestCheckPythonVersion(unittest.TestCase):
    """测试 Python 版本检查"""

    @patch("sys.stdout", new_callable=StringIO)
    def test_version_supported(self, fake_out):
        with patch.object(install.sys, "version_info", SimpleNamespace(major=3, minor=12, micro=1)):
            self.assertTrue(install.check_python_version())
        self.assertIn("[通过] Python 版本: 3.12.1", fake_out.getvalue())

    @patch("sys.stdout", new_callable=StringIO)
    def test_version_too_low(self, fake_out):
        with patch.object(install.sys, "version_info", SimpleNamespace(major=3, minor=11, micro=9)):
            self.assertFalse(install.check_python_version())
        self.assertIn("[失败] Python 版本过低: 3.11.9", fake_out.getvalue())


class TestCheckModule(unittest.TestCase):
    """测试模块检查与安装"""

    @patch("sys.stdout", new_callable=StringIO)
    def test_module_already_installed(self, fake_out):
        with patch("builtins.__import__", return_value=MagicMock()) as mock_import:
            self.assertTrue(install.check_module("pytest"))
        mock_import.assert_called_once_with("pytest")
        self.assertIn("[通过] pytest 已安装", fake_out.getvalue())

    @patch("sys.stdout", new_callable=StringIO)
    def test_module_missing_and_auto_install_success(self, fake_out):
        with patch("builtins.__import__", side_effect=ImportError), \
             patch("install.subprocess.check_call") as mock_check_call:
            self.assertTrue(install.check_module("pytest", "pytest"))
        mock_check_call.assert_called_once()
        self.assertIn("[通过] pytest 安装成功", fake_out.getvalue())

    @patch("sys.stdout", new_callable=StringIO)
    def test_module_missing_and_auto_install_failure(self, fake_out):
        with patch("builtins.__import__", side_effect=ImportError), \
             patch("install.subprocess.check_call", side_effect=install.subprocess.CalledProcessError(1, ["pip"])):
            self.assertFalse(install.check_module("pytest", "pytest"))
        self.assertIn("[失败] pytest 安装失败", fake_out.getvalue())


class TestCheckDirectoryStructure(unittest.TestCase):
    """测试目录结构检查"""

    @patch("sys.stdout", new_callable=StringIO)
    def test_all_required_files_present(self, fake_out):
        with patch("install.os.path.abspath", return_value="/repo/install.py"), \
             patch("install.os.path.exists", return_value=True), \
             patch("install.os.path.isdir", return_value=True):
            self.assertTrue(install.check_directory_structure())
        output = fake_out.getvalue()
        self.assertIn("[通过] main.py", output)
        self.assertIn("[通过] results/ 目录存在", output)

    @patch("sys.stdout", new_callable=StringIO)
    def test_missing_required_file_marks_failure(self, fake_out):
        def exists_side_effect(path):
            return not path.endswith("cli.py")

        with patch("install.os.path.abspath", return_value="/repo/install.py"), \
             patch("install.os.path.exists", side_effect=exists_side_effect), \
             patch("install.os.path.isdir", return_value=False):
            self.assertFalse(install.check_directory_structure())
        output = fake_out.getvalue()
        self.assertIn("[缺失] cli.py", output)
        self.assertIn("[提示] results/ 目录不存在", output)


class TestInstallMain(unittest.TestCase):
    """测试 install.main 汇总行为"""

    @patch("sys.stdout", new_callable=StringIO)
    def test_main_returns_success(self, fake_out):
        with patch("install.check_python_version", return_value=True), \
             patch("install.check_module", return_value=True), \
             patch("install.check_directory_structure", return_value=True):
            self.assertEqual(install.main(), 0)
        self.assertIn("环境检查全部通过！可以开始使用。", fake_out.getvalue())

    @patch("sys.stdout", new_callable=StringIO)
    def test_main_returns_failure(self, fake_out):
        with patch("install.check_python_version", return_value=False), \
             patch("install.check_module", return_value=True), \
             patch("install.check_directory_structure", return_value=True):
            self.assertEqual(install.main(), 1)
        self.assertIn("环境检查未通过", fake_out.getvalue())


if __name__ == "__main__":
    unittest.main()

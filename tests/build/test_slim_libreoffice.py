from __future__ import annotations

from pathlib import Path
import sys
import pytest

# 将项目根目录与 build_scripts 目录加入导入路径
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "build_scripts"))

from slim_libreoffice import (
    PRUNE_RULES,
    execute,
    main as slim_main,
    prune_plan,
)


def create_fake_lo_tree(root: Path) -> dict[str, Path]:
    """构造用于测试的假 LibreOffice 目录树。"""
    files = {
        # program/ 核心目录（绝对不可删）
        "program_soffice_exe": root / "program" / "soffice.exe",
        "program_soffice_bin": root / "program" / "soffice.bin",
        "program_soffice_com": root / "program" / "soffice.com",
        "program_writer_dll": root / "program" / "writer.dll",
        # share/program/ （必须绝对保护）
        "share_program_dll": root / "share" / "program" / "keep.dll",
        # share/help/ （应该被删除）
        "share_help_doc": root / "share" / "help" / "zh-CN" / "index.html",
        # share/template/ (应该被删除)
        "share_template_ott": root / "share" / "template" / "common" / "letter.ott",
        "share_templates_ott": root / "share" / "templates" / "modern.ott",
        # share/extensions/ (应该被删除)
        "share_ext_oxt": root / "share" / "extensions" / "sample" / "ext.oxt",
        # share/dict/ 中的保留项（en*, zh*）
        "share_dict_en_dat": root / "share" / "dict" / "en_US.dat",
        "share_dict_zh_dat": root / "share" / "dict" / "zh_CN.dat",
        "share_dict_en_aff": root / "share" / "dict" / "en_US.aff",
        "share_dict_zh_dic": root / "share" / "dict" / "zh_CN.dic",
        # share/dict/ 中的待删除项（非 en, zh）
        "share_dict_fr_dat": root / "share" / "dict" / "fr_FR.dat",
        "share_dict_de_dat": root / "share" / "dict" / "de_DE.dat",
        "share_dict_es_dic": root / "share" / "dict" / "es_ES.dic",
        # share/ 其他配置/注册表项（必须保留，不可误删）
        "share_config_cfg": root / "share" / "config" / "soffice.cfg" / "swriter.xml",
        "share_registry_xcd": root / "share" / "registry" / "main.xcd",
    }
    for p in files.values():
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"dummy-content-123")
    return files


def test_prune_plan_targets_and_preservation(tmp_path: Path):
    """验证 prune_plan 命中预期黑名单，且不碰 program/ 与 share/program/。"""
    fake_lo = tmp_path / "fake_lo"
    create_fake_lo_tree(fake_lo)

    plan = prune_plan(fake_lo)
    plan_posix = {p.as_posix() for p in plan}

    # 1. 绝对保护 program/ 及其子项
    assert not any(p.parts[0] == "program" for p in plan)

    # 2. 绝对保护 share/program/ 及其子项
    assert not any(p.parts[:2] == ("share", "program") for p in plan)

    # 3. 绝对保护 share/config 与 share/registry
    assert not any("config" in p.parts for p in plan)
    assert not any("registry" in p.parts for p in plan)

    # 4. 命中 share/help、share/template、share/extensions
    assert "share/help/zh-CN/index.html" in plan_posix
    assert "share/template/common/letter.ott" in plan_posix
    assert "share/templates/modern.ott" in plan_posix
    assert "share/extensions/sample/ext.oxt" in plan_posix

    # 5. share/dict/ 中精准保留 en/zh 项，移除非 en/zh 项
    assert "share/dict/fr_FR.dat" in plan_posix
    assert "share/dict/de_DE.dat" in plan_posix
    assert "share/dict/es_ES.dic" in plan_posix

    assert "share/dict/en_US.dat" not in plan_posix
    assert "share/dict/zh_CN.dat" not in plan_posix
    assert "share/dict/en_US.aff" not in plan_posix
    assert "share/dict/zh_CN.dic" not in plan_posix


def test_cli_dry_run_does_not_modify_filesystem(tmp_path: Path, capsys):
    """验证 --dry-run 仅输出规划清单，完全不修改源目录且不创建输出目录。"""
    fake_lo = tmp_path / "fake_lo"
    create_fake_lo_tree(fake_lo)
    output_dir = tmp_path / "output_runtime"

    ret = slim_main(["--source", str(fake_lo), "--output", str(output_dir), "--dry-run"])
    assert ret == 0

    # 验证输出目录未被创建
    assert not output_dir.exists()

    # 验证输出信息包含 [dry-run]
    captured = capsys.readouterr()
    assert "[dry-run]" in captured.out
    assert "share/help" in captured.out
    assert "share/dict/fr_FR.dat" in captured.out
    assert "program/soffice" not in captured.out


def test_execute_prunes_correctly_without_smoke(tmp_path: Path):
    """验证 execute 在 --no-smoke 模式下的实际复制与删除效果。"""
    fake_lo = tmp_path / "fake_lo"
    create_fake_lo_tree(fake_lo)
    output_dir = tmp_path / "output_runtime"

    plan = execute(source=fake_lo, output=output_dir, smoke=False)
    assert len(plan) > 0

    # 验证 output 存在且核心保留
    assert (output_dir / "program" / "soffice.exe").is_file()
    assert (output_dir / "share" / "program" / "keep.dll").is_file()
    assert (output_dir / "share" / "dict" / "zh_CN.dat").is_file()
    assert (output_dir / "share" / "dict" / "en_US.dat").is_file()
    assert (output_dir / "share" / "config" / "soffice.cfg" / "swriter.xml").is_file()

    # 验证 output 中黑名单目标已被清除
    assert not (output_dir / "share" / "dict" / "fr_FR.dat").exists()
    assert not (output_dir / "share" / "dict" / "de_DE.dat").exists()
    assert not (output_dir / "share" / "help" / "zh-CN" / "index.html").exists()
    assert not (output_dir / "share" / "template" / "common" / "letter.ott").exists()


def test_smoke_failure_triggers_rollback(tmp_path: Path, monkeypatch):
    """验证当冒烟测试失败时，会回滚删除的文件并抛出异常。"""
    fake_lo = tmp_path / "fake_lo"
    create_fake_lo_tree(fake_lo)
    output_dir = tmp_path / "output_runtime"

    # mock run_smoke_test 返回 False 模拟转换崩溃或失败（现仅 output 副本一次冒烟）
    import slim_libreoffice

    monkeypatch.setattr(slim_libreoffice, "run_smoke_test", lambda *a, **k: False)

    # 伪造一个可执行的 soffice 文件供探测
    fake_soffice = output_dir / "program" / "soffice.com"

    with pytest.raises(RuntimeError, match="冒烟验证失败，已回滚"):
        execute(source=fake_lo, output=output_dir, soffice=fake_soffice, smoke=True)

    # 验证回滚后，原本应该被删除的文件依然被恢复保存在 output_dir 中
    assert (output_dir / "share" / "dict" / "fr_FR.dat").is_file()
    assert (output_dir / "share" / "help" / "zh-CN" / "index.html").is_file()
    assert (output_dir / "program" / "soffice.exe").is_file()

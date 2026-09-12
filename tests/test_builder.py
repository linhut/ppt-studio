# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""Builder 黄金结构单测：对 json2pptx 完整链路做语义断言（页数/文本/图表/表格/复合组件），
以及新增工具脚本（validate / md2deck / pptx2json / new_project / wizard_interactive）的集成验收。"""

import json

import pytest

import json2pptx
import validate as validate_mod
import md2deck
import pptx2json
import new_project
import wizard_interactive


def _deck():
    return {
        "canvas": {"w": 960, "h": 540},
        "theme": {"colors": {
            "$bg": "#FFFFFF", "$paper": "#F5F8FC", "$primary": "#1A3C8B",
            "$accent": "#E67733", "$ink": "#1A3C8B", "$text": "#222222",
            "$muted": "#666666", "$line": "#D0D8E4", "$green": "#188050",
            "$brass": "#E67733",
        }},
        "slides": [
            {"id": "01", "title": "封面", "background": {"color": "$bg"},
             "elements": [
                 {"type": "text", "role": "title", "x": 60, "y": 130, "w": 840,
                  "h": 90, "text": "黄金测试标题", "fontSize": 54,
                  "color": "$primary", "bold": True, "align": "center"},
                 {"type": "shape", "shapeName": "rect", "x": 0, "y": 0, "w": 960,
                  "h": 8, "fill": {"color": "$primary"}},
             ]},
            {"id": "02", "title": "KPI 页",
             "elements": [
                 {"type": "text", "role": "title", "x": 60, "y": 20, "w": 840,
                  "h": 44, "text": "关键指标", "fontSize": 38,
                  "color": "$primary", "bold": True},
                 {"type": "cards_1x4_info", "x": 60, "y": 120, "w": 840,
                  "h": 140, "items": [
                      {"num": "99%", "label": "满意度"},
                      {"num": "1.2k", "label": "用户数"},
                      {"num": "36", "label": "份额"},
                      {"num": "5★", "label": "评分"}]},
             ]},
            {"id": "03", "title": "数据页",
             "elements": [
                 {"type": "chart", "x": 60, "y": 100, "w": 520, "h": 360,
                  "color": "$green",
                  "data": {"rows": [["Q1", "单位", 42], ["Q2", "单位", 55]]}},
                 {"type": "table", "x": 610, "y": 100, "w": 290, "h": 200,
                  "rows": 3, "cols": 2, "header_color": "$ink",
                  "data": [["", "方案A"], ["成本", "低"], ["效率", "高"]]},
             ]},
            {"id": "04", "title": "流程",
             "elements": [
                 {"type": "process_steps", "x": 48, "y": 130, "w": 864,
                  "h": 300, "color": "$accent",
                  "steps": [{"title": "步骤1", "desc": "说明一"},
                            {"title": "步骤2", "desc": "说明二"}]},
             ]},
        ],
    }


def _build(data, tmp_path, no_fade=False):
    src = tmp_path / "deck.json"
    src.write_text(json.dumps(data), encoding="utf-8")
    b = json2pptx.Builder(data, tmp_path)
    warnings = b.build(no_fade=no_fade)
    return b.prs, warnings


class TestBuilder:
    def test_slide_count(self, tmp_path):
        prs, _ = _build(_deck(), tmp_path)
        assert len(prs.slides._sldIdLst) == 4

    def test_title_text_present(self, tmp_path):
        prs, _ = _build(_deck(), tmp_path)
        texts = []
        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
        assert any("黄金测试标题" in t for t in texts)

    def test_kpi_composite_renders(self, tmp_path):
        prs, _ = _build(_deck(), tmp_path)
        # cards_1x4_info 至少渲染出 4 个卡片文本（num/label 各至少一帧文本）
        texts = []
        for shape in prs.slides[1].shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
        joined = "\n".join(texts)
        assert "99%" in joined and "满意度" in joined

    def test_chart_and_table(self, tmp_path):
        prs, _ = _build(_deck(), tmp_path)
        charts = [s for s in prs.slides[2].shapes if getattr(s, "has_chart", False)]
        tables = [s for s in prs.slides[2].shapes if getattr(s, "has_table", False)]
        assert len(charts) == 1
        assert len(tables) == 1
        cats = [str(c) for c in charts[0].chart.plots[0].categories]
        assert cats == ["Q1", "Q2"]
        cells = [c.text for row in tables[0].table.rows for c in row.cells]
        assert "方案A" in cells and "成本" in cells

    def test_process_steps(self, tmp_path):
        prs, _ = _build(_deck(), tmp_path)
        texts = [s.text_frame.text for s in prs.slides[3].shapes
                 if s.has_text_frame]
        assert any("步骤1" in t for t in texts)
        assert any("步骤2" in t for t in texts)

    def test_no_warnings_on_golden_deck(self, tmp_path):
        _, warnings = _build(_deck(), tmp_path)
        assert warnings == []

    def test_save_produces_valid_pptx(self, tmp_path):
        data = _deck()
        b = json2pptx.Builder(data, tmp_path)
        b.build(no_fade=True)
        out = tmp_path / "out.pptx"
        b.prs.save(str(out))
        assert out.is_file() and out.stat().st_size > 0


class TestValidate:
    def test_valid_deck_passes(self, tmp_path):
        src = tmp_path / "d.json"
        src.write_text(json.dumps(_deck()), encoding="utf-8")
        data = json.loads(src.read_text(encoding="utf-8"))
        issues = validate_mod.validate_deck(data, {}, src.read_text(encoding="utf-8"))
        assert not [i for i in issues if i[0] == "error"]

    def test_broken_deck_reports_errors(self, tmp_path):
        d = _deck()
        d["slides"][0]["elements"][0]["type"] = "hologram"
        d["slides"][1]["id"] = d["slides"][0]["id"]
        src = tmp_path / "d.json"
        src.write_text(json.dumps(d), encoding="utf-8")
        text = src.read_text(encoding="utf-8")
        issues = validate_mod.validate_deck(json.loads(text), validate_mod.build_offset_map(text), text)
        errs = [i for i in issues if i[0] == "error"]
        kinds = {i[2].split()[-1] for i in errs}
        assert any("hologram" in i[2] for i in errs)
        assert "重复" in " ".join(i[2] for i in errs)


class TestMd2deck:
    def test_parse_and_build(self, tmp_path):
        md = "# 汇报\n> 副标题\n\n## 章节一\n- 要点A：说明A\n- 要点B\n\n## 指标\n- 99%\n- 1.2k\n"
        title, subtitle, slides = md2deck.parse_markdown(md)
        assert title == "汇报" and subtitle == "副标题"
        assert len(slides) == 2
        deck = md2deck.build_deck(title, subtitle, slides, {})
        assert len(deck["slides"]) == 3  # 封面 + 2 页

    def test_kpi_heuristic(self, tmp_path):
        md = "## 指标\n- 99.9%：可用性\n- 1.2k：用户\n"
        _, _, slides = md2deck.parse_markdown(md)
        deck = md2deck.build_deck("t", "", slides, {})
        # 数字型条目应走 kpi 布局（cards_1x4_info）
        types = [e["type"] for e in deck["slides"][1]["elements"]]
        assert "cards_1x4_info" in types


class TestPptx2json:
    def test_roundtrip(self, tmp_path):
        b = json2pptx.Builder(_deck(), tmp_path)
        b.build(no_fade=True)
        pptx = tmp_path / "src.pptx"
        b.prs.save(str(pptx))
        deck, warnings = pptx2json.convert(pptx)
        assert len(deck["slides"]) == 4
        # 回环产物可通过 schema 校验
        import validate as v
        issues = v.validate_deck(deck, {}, json.dumps(deck))
        assert not [i for i in issues if i[0] == "error"]


class TestNewProject:
    def test_scaffold_and_count(self, tmp_path):
        rc = new_project.build_new("my-deck", "tech", 8, str(tmp_path))
        assert rc == 0
        out = tmp_path / "my-deck" / "deck.json"
        assert out.is_file()
        deck = json.loads(out.read_text(encoding="utf-8"))
        assert len(deck["slides"]) == 8

    def test_bad_theme_fails(self, tmp_path):
        assert new_project.build_new("d", "no-such-theme", 8, str(tmp_path)) == 2


class TestWizard:
    def test_defaults_build(self, tmp_path):
        answers = wizard_interactive.collect_answers(defaults_only=True)
        out = tmp_path / "w.json"
        assert wizard_interactive.build(answers, out) == 0
        deck = json.loads(out.read_text(encoding="utf-8"))
        assert len(deck["slides"]) >= 4


class TestEngineExtensions:
    """引擎扩展：canvas 尺寸 / 字体配置 / 图表类型。"""

    def _deck(self):
        return _deck()

    def test_canvas_4x3(self, tmp_path):
        d = _deck()
        d["canvas"] = {"w": 960, "h": 720}  # 4:3
        b = json2pptx.Builder(d, tmp_path)
        b.build(no_fade=True)
        assert b.prs.slide_width == 960 * 12700
        assert b.prs.slide_height == 720 * 12700

    def test_fonts_config(self, tmp_path):
        d = _deck()
        d["theme"]["fonts"] = {"heading": "Arial Black", "body": "SimSun",
                               "stat": "Consolas", "table": "KaiTi"}
        b = json2pptx.Builder(d, tmp_path)
        b.build(no_fade=True)
        # 标题角色 → heading 字体
        title_run = None
        for shape in b.prs.slides[0].shapes:
            if shape.has_text_frame and "黄金测试标题" in shape.text_frame.text:
                title_run = shape.text_frame.paragraphs[0].runs[0]
        assert title_run is not None and title_run.font.name == "Arial Black"

    def test_chart_legacy_unit_column(self, tmp_path):
        """兼容 [分类, 单位, 值] 旧格式：单位列不产生虚假系列。"""
        d = _deck()
        b = json2pptx.Builder(d, tmp_path)
        b.build(no_fade=True)
        charts = [s for s in b.prs.slides[2].shapes if getattr(s, "has_chart", False)]
        assert len(charts) == 1
        assert len(charts[0].chart.plots[0].series) == 1

    def test_chart_multi_series(self, tmp_path):
        d = _deck()
        d["slides"][2]["elements"][0]["data"]["rows"] = [["A", 10, 20], ["B", 30, 40]]
        b = json2pptx.Builder(d, tmp_path)
        b.build(no_fade=True)
        charts = [s for s in b.prs.slides[2].shapes if getattr(s, "has_chart", False)]
        assert len(charts) == 1
        assert len(charts[0].chart.plots[0].series) == 2

    def test_chart_types_build(self, tmp_path):
        for ctype in ("line", "pie", "radar", "stacked"):
            d = _deck()
            d["slides"][2]["elements"][0]["chartType"] = ctype
            b = json2pptx.Builder(d, tmp_path)
            b.build(no_fade=True)
            charts = [s for s in b.prs.slides[2].shapes if getattr(s, "has_chart", False)]
            assert len(charts) == 1, ctype

    def test_chart_scatter(self, tmp_path):
        d = _deck()
        d["slides"][2]["elements"][0]["chartType"] = "scatter"
        d["slides"][2]["elements"][0]["data"]["rows"] = [["p1", 1, 2], ["p2", 3, 5]]
        b = json2pptx.Builder(d, tmp_path)
        b.build(no_fade=True)
        charts = [s for s in b.prs.slides[2].shapes if getattr(s, "has_chart", False)]
        assert len(charts) == 1
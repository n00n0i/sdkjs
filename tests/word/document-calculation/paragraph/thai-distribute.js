/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	const charWidth = AscTest.CharWidth * AscTest.FontSize;

	let dc = new AscWord.CDocumentContent();
	dc.ClearContent(false);

	let para = new AscWord.Paragraph();
	dc.AddToContent(0, para);

	let run = new AscWord.CRun();
	para.AddToContent(0, run);

	function recalculate(width)
	{
		dc.Reset(0, 0, width, 10000);
		dc.Recalculate_Page(0, true);
	}

	QUnit.module("Thai distribute layout");

	QUnit.test("ดูการบ้าน breaks after ดู, not as ดูกา|รบ้าน", function (assert)
	{
		run.ClearContent();
		run.AddText("ดูการบ้าน");
		para.Set_Align(AscCommon.align_ThaiDistributed);
		// "การบ้าน" is 7 code points. The whole string is 9. This width fits the
		// second word and does not fit the unsegmented run.
		recalculate(charWidth * 7.5);

		assert.strictEqual(para.GetLinesCount(), 2, "two lines");
		assert.strictEqual(para.GetTextOnLine(0), "ดู", "first line is the word ดู");
		assert.strictEqual(para.GetTextOnLine(1), "การบ้าน", "second line is การบ้าน");
	});
});

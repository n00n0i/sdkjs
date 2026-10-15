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

(function (root)
{
	var AscWord = root.AscWord || (root.window && root.window.AscWord);

	function near(a, b)
	{
		return Math.abs(a - b) < 1e-9;
	}

	function sameList(actual, expected)
	{
		if (!actual || actual.length !== expected.length)
			return false;
		for (var i = 0; i < expected.length; i++)
		{
			if (actual[i] !== expected[i])
				return false;
		}
		return true;
	}

	function layoutAdvances(widths, marks, extra, spaces, stretch)
	{
		var clusters = 0;
		for (var i = 0; i < marks.length; i++)
		{
			if (!marks[i])
				clusters++;
		}
		var gaps = AscWord.calcDistributeGaps(extra, clusters, spaces, stretch);
		var out = [];
		for (var n = 0; n < widths.length; n++)
		{
			var nextIsMark = n + 1 < marks.length && marks[n + 1];
			var w = widths[n];
			if (AscWord.clusterTakesJustify(nextIsMark))
				w += gaps.justifyWord;
			out.push(w);
		}
		return { clusters: clusters, gaps: gaps, widths: out };
	}

	function collectFailures()
	{
		var failures = [];
		function check(ok, name, detail)
		{
			if (!ok)
				failures.push(name + (detail ? ": " + detail : ""));
		}

		var both = AscWord.x2tJcToByte("both");
		var distribute = AscWord.x2tJcToByte("distribute");
		var thai = AscWord.x2tJcToByte("thaiDistribute");
		check(both === 3 && distribute === 4 && thai === 6, "x2t bytes", [both, distribute, thai].join(","));
		check(both !== distribute && distribute !== thai && both !== thai, "three jc values stay distinct");
		check(AscWord.x2tByteToJc(both) === "both", "byte 3 reads back as both");
		check(AscWord.x2tByteToJc(distribute) === "distribute", "byte 4 reads back as distribute");
		check(AscWord.x2tByteToJc(thai) === "thaiDistribute", "byte 6 reads back as thaiDistribute");
		check(AscWord.x2tSetValueFromByte(both) === 0, "both -> EJc both (0)");
		check(AscWord.x2tSetValueFromByte(distribute) === 2, "distribute -> EJc distribute (2)");
		check(AscWord.x2tSetValueFromByte(thai) === 9, "thaiDistribute -> EJc thaiDistribute (9)");

		check(AscWord.alignFromWordJc("both") === both, "json both");
		check(AscWord.alignFromWordJc("distribute") === distribute, "json distribute");
		check(AscWord.alignFromWordJc("thaiDistribute") === thai, "json thaiDistribute");
		check(AscWord.alignFromWordJc("thaiDist") === thai, "json thaiDist");
		check(AscWord.alignToWordJc(thai) === "thaiDistribute", "json write thaiDistribute");
		check(AscWord.alignToWordJc(distribute) === "distribute", "json write distribute");
		check(AscWord.alignToWordJc(both) === "both", "json write both");
		check(AscWord.alignFromDrawingAlgn("dist") === distribute, "drawing dist");
		check(AscWord.alignFromDrawingAlgn("thaiDist") === thai, "drawing thaiDist");
		check(AscWord.alignToDrawingAlgn(distribute) === "dist", "drawing write dist");
		check(AscWord.alignToDrawingAlgn(thai) === "thaiDist", "drawing write thaiDist");
		check(AscWord.paraAlignFromBinaryByte(thai) === thai && AscWord.paraAlignToBinaryByte(thai) === thai, "binary byte is the align value");

		var thaiWord = "ดูการบ้าน";
		check(sameList(AscWord.dictionaryBreakAfterIndices(thaiWord), [1]), "dictionary ดูการบ้าน breaks after ดู", JSON.stringify(AscWord.dictionaryBreakAfterIndices(thaiWord)));
		check(sameList(AscWord.thaiBreakAfterIndices(thaiWord), [1]), "segmenter-or-dictionary ดูการบ้าน", JSON.stringify(AscWord.thaiBreakAfterIndices(thaiWord)));
		check(AscWord.dictionaryBreakAfterIndices(thaiWord).indexOf(3) === -1, "does not break as ดูกา|รบ้าน");
		check(sameList(AscWord.dictionaryBreakAfterIndices("hello"), []), "no breaks in hello");
		check(sameList(AscWord.dictionaryBreakAfterIndices("hello world"), []), "no breaks in hello world");
		check(sameList(AscWord.thaiBreakAfterIndices("hello world"), []), "segmenter leaves hello world");

		var mixed = AscWord.dictionaryBreakAfterIndices("abcดูการบ้าน");
		check(sameList(mixed, [4]), "mixed abcดูการบ้าน breaks after ดู", JSON.stringify(mixed));

		var cluster = AscWord.dictionaryBreakAfterIndices("กิข");
		check(sameList(cluster, [1]), "กิข breaks after the mark", JSON.stringify(cluster));
		check(cluster.indexOf(0) === -1, "no break between ก and ิ");

		check(AscWord.countGraphemeClusters([0x0E01, 0x0E34, 0x0E02]) === 2, "กิข is two clusters");
		check(AscWord.clusterTakesJustify(true) === false, "no gap before a combining mark");
		check(AscWord.clusterTakesJustify(false) === true, "cluster terminator takes a gap");

		var spread = AscWord.calcDistributeGaps(12, 4, 2, true);
		check(spread.gaps === 5 && near(spread.justifyWord, 2.4) && near(spread.justifySpace, 2.4), "12 extra across 4 clusters and 2 spaces", JSON.stringify(spread));
		var stopped = AscWord.calcDistributeGaps(12, 4, 2, false);
		check(stopped.gaps === 0 && stopped.justifyWord === 0 && stopped.justifySpace === 0, "no stretch");
		var empty = AscWord.calcDistributeGaps(0, 4, 2, true);
		check(empty.justifyWord === 0 && empty.gaps === 0, "no extra");
		var single = AscWord.calcDistributeGaps(12, 1, 0, true);
		check(single.gaps === 0 && single.justifyWord === 0, "one cluster has no gap");

		var laid = layoutAdvances([10, 0, 10], [false, true, false], 8, 0, true);
		check(laid.clusters === 2 && laid.gaps.gaps === 1 && near(laid.gaps.justifyWord, 8), "two clusters, one gap", JSON.stringify(laid.gaps));
		check(near(laid.widths[0], 10) && near(laid.widths[1], 8) && near(laid.widths[2], 18), "mark keeps the gap, last cluster keeps trailing advance", JSON.stringify(laid.widths));

		return failures;
	}

	root.AscTestThaiDistribute = { run: collectFailures };

	if (typeof QUnit !== "undefined")
	{
		QUnit.module("Thai distribute");
		QUnit.test("jc round-trip, clusters, and ดูการบ้าน breaks", function (assert)
		{
			var failures = collectFailures();
			assert.strictEqual(failures.length, 0, failures.join("\n"));
		});
	}
})(typeof window !== "undefined" ? window : globalThis);

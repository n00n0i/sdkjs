/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window)
{
	var AscWord = window.AscWord = window.AscWord || {};

	// Binary paragraph jc bytes. These are the sdkjs align constants and the
	// bytes x2t writes after the 9.4.0 thai-distribute change in ONLYOFFICE/core
	// (OOXML/Binary/Document BinWriter + BinReader). Do not renumber them:
	// 5 is align_CenterContinuous.
	var BIN_ALIGN_RIGHT = 0;
	var BIN_ALIGN_LEFT = 1;
	var BIN_ALIGN_CENTER = 2;
	var BIN_ALIGN_JUSTIFY = 3;
	var BIN_ALIGN_DISTRIBUTED = 4;
	var BIN_ALIGN_THAI_DISTRIBUTED = 6;

	// EJc values in OOXML/Common/SimpleTypes_Word.h (core v9.4.0.129).
	var EJC_BOTH = 0;
	var EJC_DISTRIBUTE = 2;
	var EJC_THAI_DISTRIBUTE = 9;

	function isThaiCombiningMark(nCode)
	{
		return nCode === 0x0E31
			|| (nCode >= 0x0E34 && nCode <= 0x0E3A)
			|| (nCode >= 0x0E47 && nCode <= 0x0E4E);
	}

	function isThaiScript(nCode)
	{
		return nCode >= 0x0E00 && nCode <= 0x0E7F;
	}

	function textHasThai(sText)
	{
		for (var i = 0; i < sText.length; i++)
		{
			if (isThaiScript(sText.charCodeAt(i)))
				return true;
		}
		return false;
	}

	var thaiTrie = null;

	function buildThaiTrie(sList)
	{
		var oRoot = { $: 0 };
		if (!sList)
			return oRoot;
		var arr = sList.split("\n");
		for (var nWord = 0; nWord < arr.length; nWord++)
		{
			var sWord = arr[nWord];
			if (!sWord)
				continue;
			var oNode = oRoot;
			for (var i = 0; i < sWord.length; i++)
			{
				var nCode = sWord.charCodeAt(i);
				var oNext = oNode[nCode];
				if (!oNext)
				{
					oNext = { $: 0 };
					oNode[nCode] = oNext;
				}
				oNode = oNext;
			}
			oNode.$ = 1;
		}
		return oRoot;
	}

	function getThaiTrie()
	{
		if (!thaiTrie)
			thaiTrie = buildThaiTrie(AscWord.THAI_WORD_LIST || "");
		return thaiTrie;
	}

	var segmenterState = 0;
	var thaiSegmenter = null;

	function getThaiSegmenter()
	{
		if (segmenterState !== 0)
			return thaiSegmenter;
		segmenterState = -1;
		thaiSegmenter = null;
		try
		{
			if (typeof Intl !== "undefined" && typeof Intl.Segmenter === "function")
			{
				var oSeg = new Intl.Segmenter("th", { granularity: "word" });
				var arr = [];
				var oIter = oSeg.segment("ดูการบ้าน");
				if (oIter && typeof oIter[Symbol.iterator] === "function")
				{
					var it = oIter[Symbol.iterator]();
					var step;
					while (!(step = it.next()).done)
						arr.push(step.value);
				}
				var sFirst = "";
				for (var i = 0; i < arr.length; i++)
				{
					if (arr[i] && arr[i].isWordLike)
					{
						sFirst = arr[i].segment;
						break;
					}
				}
				if (sFirst === "ดู" && arr.length >= 2)
				{
					thaiSegmenter = oSeg;
					segmenterState = 1;
				}
			}
		}
		catch (e)
		{
			thaiSegmenter = null;
			segmenterState = -1;
		}
		return thaiSegmenter;
	}

	function pushBreak(arrBreaks, nIndex, nLen)
	{
		if (nIndex >= 0 && nIndex < nLen - 1)
			arrBreaks.push(nIndex);
	}

	function dictionaryBreaks(sText)
	{
		var oTrie = getThaiTrie();
		var arrBreaks = [];
		var nLen = sText.length;
		var i = 0;
		while (i < nLen)
		{
			var nCode = sText.charCodeAt(i);
			if (!isThaiScript(nCode))
			{
				i++;
				continue;
			}
			if (isThaiCombiningMark(nCode))
			{
				i++;
				continue;
			}

			var oNode = oTrie;
			var nBest = i;
			var j = i;
			while (j < nLen)
			{
				var nChar = sText.charCodeAt(j);
				if (!oNode[nChar])
					break;
				oNode = oNode[nChar];
				j++;
				if (oNode.$)
					nBest = j;
			}

			var nEnd;
			if (nBest > i)
				nEnd = nBest;
			else
				nEnd = i + 1;

			while (nEnd < nLen && isThaiCombiningMark(sText.charCodeAt(nEnd)))
				nEnd++;

			pushBreak(arrBreaks, nEnd - 1, nLen);
			i = nEnd;
		}
		return arrBreaks;
	}

	function segmenterBreaks(sText, oSeg)
	{
		var arrBreaks = [];
		var nLen = sText.length;
		var oIter = oSeg.segment(sText);
		var it = oIter[Symbol.iterator]();
		var step;
		while (!(step = it.next()).done)
		{
			var oPart = step.value;
			if (!oPart || !oPart.isWordLike || !oPart.segment)
				continue;
			var nEnd = oPart.index + oPart.segment.length;
			while (nEnd < nLen && isThaiCombiningMark(sText.charCodeAt(nEnd)))
				nEnd++;
			// Only emit a break when the segment contains Thai. Latin words
			// already break on spaces, which are separate run items.
			var bThai = false;
			for (var i = oPart.index; i < nEnd; i++)
			{
				if (isThaiScript(sText.charCodeAt(i)))
				{
					bThai = true;
					break;
				}
			}
			if (bThai)
				pushBreak(arrBreaks, nEnd - 1, nLen);
		}
		return arrBreaks;
	}

	function thaiBreakAfterIndices(sText)
	{
		if (!sText || !textHasThai(sText))
			return [];
		var oSeg = getThaiSegmenter();
		if (oSeg)
			return segmenterBreaks(sText, oSeg);
		return dictionaryBreaks(sText);
	}

	function dictionaryBreakAfterIndices(sText)
	{
		if (!sText || !textHasThai(sText))
			return [];
		return dictionaryBreaks(sText);
	}

	/**
	 * Extra width for thaiDistribute / distribute.
	 * Slots are the gaps between grapheme clusters, plus each counted space.
	 * justifyWord is added to every cluster terminator, including the last one
	 * (that last addition is trailing advance, so the ink ends on the margin).
	 * @returns {{justifyWord: number, justifySpace: number, gaps: number}}
	 */
	function calcDistributeGaps(nExtra, nClusters, nSpaces, bStretch)
	{
		if (!bStretch || !(nExtra > 0.001))
			return { justifyWord: 0, justifySpace: 0, gaps: 0 };
		var nBetween = nClusters > 0 ? nClusters - 1 : 0;
		var nSpacesCount = nSpaces > 0 ? nSpaces : 0;
		var nGaps = nBetween + nSpacesCount;
		if (nGaps <= 0)
			return { justifyWord: 0, justifySpace: 0, gaps: 0 };
		var nPer = nExtra / nGaps;
		return { justifyWord: nPer, justifySpace: nPer, gaps: nGaps };
	}

	function clusterTakesJustify(bNextIsCombiningMark)
	{
		return !bNextIsCombiningMark;
	}

	function countGraphemeClusters(arrCodes)
	{
		var nClusters = 0;
		for (var i = 0; i < arrCodes.length; i++)
		{
			if (!isThaiCombiningMark(arrCodes[i]))
				nClusters++;
		}
		return nClusters;
	}

	function alignFromWordJc(sJc)
	{
		switch (sJc)
		{
			case "end":
			case "right":
				return BIN_ALIGN_RIGHT;
			case "start":
			case "left":
				return BIN_ALIGN_LEFT;
			case "center":
				return BIN_ALIGN_CENTER;
			case "both":
			case "justify":
				return BIN_ALIGN_JUSTIFY;
			case "distribute":
				return BIN_ALIGN_DISTRIBUTED;
			case "thaiDistribute":
			case "thaiDist":
				return BIN_ALIGN_THAI_DISTRIBUTED;
			default:
				return undefined;
		}
	}

	function alignToWordJc(nAlign)
	{
		switch (nAlign)
		{
			case BIN_ALIGN_RIGHT:
				return "end";
			case BIN_ALIGN_LEFT:
				return "start";
			case BIN_ALIGN_CENTER:
				return "center";
			case BIN_ALIGN_JUSTIFY:
				return "both";
			case BIN_ALIGN_DISTRIBUTED:
				return "distribute";
			case BIN_ALIGN_THAI_DISTRIBUTED:
				return "thaiDistribute";
			default:
				return undefined;
		}
	}

	function alignFromDrawingAlgn(sAlgn)
	{
		switch (sAlgn)
		{
			case "ctr":
				return BIN_ALIGN_CENTER;
			case "dist":
				return BIN_ALIGN_DISTRIBUTED;
			case "just":
			case "justLow":
				return BIN_ALIGN_JUSTIFY;
			case "l":
				return BIN_ALIGN_LEFT;
			case "r":
				return BIN_ALIGN_RIGHT;
			case "thaiDist":
				return BIN_ALIGN_THAI_DISTRIBUTED;
			default:
				return undefined;
		}
	}

	function alignToDrawingAlgn(nAlign)
	{
		switch (nAlign)
		{
			case BIN_ALIGN_CENTER:
				return "ctr";
			case BIN_ALIGN_JUSTIFY:
				return "just";
			case BIN_ALIGN_LEFT:
				return "l";
			case BIN_ALIGN_RIGHT:
				return "r";
			case BIN_ALIGN_DISTRIBUTED:
				return "dist";
			case BIN_ALIGN_THAI_DISTRIBUTED:
				return "thaiDist";
			default:
				return undefined;
		}
	}

	function paraAlignFromBinaryByte(nByte)
	{
		return nByte;
	}

	function paraAlignToBinaryByte(nAlign)
	{
		return nAlign;
	}

	/**
	 * x2t docx->bin byte for a paragraph w:jc value, after the core patch.
	 * Unpatched x2t writes 3 (align_Justify) for both, distribute and thaiDistribute.
	 */
	function x2tJcToByte(sJc)
	{
		switch (sJc)
		{
			case "center":
				return BIN_ALIGN_CENTER;
			case "left":
			case "start":
				return BIN_ALIGN_LEFT;
			case "right":
			case "end":
				return BIN_ALIGN_RIGHT;
			case "both":
				return BIN_ALIGN_JUSTIFY;
			case "distribute":
				return BIN_ALIGN_DISTRIBUTED;
			case "thaiDistribute":
				return BIN_ALIGN_THAI_DISTRIBUTED;
			default:
				return BIN_ALIGN_LEFT;
		}
	}

	/**
	 * x2t bin->docx w:jc val for a paragraph alignment byte, after the core patch.
	 * Mirrors Binary_pPrReader SetValueFromByte + CJc::ToString.
	 */
	function x2tByteToJc(nByte)
	{
		switch (nByte)
		{
			case BIN_ALIGN_RIGHT:
				return "right";
			case BIN_ALIGN_LEFT:
				return "left";
			case BIN_ALIGN_CENTER:
				return "center";
			case BIN_ALIGN_JUSTIFY:
				return "both";
			case BIN_ALIGN_DISTRIBUTED:
				return "distribute";
			case BIN_ALIGN_THAI_DISTRIBUTED:
				return "thaiDistribute";
			default:
				return null;
		}
	}

	function x2tSetValueFromByte(nByte)
	{
		switch (nByte)
		{
			case BIN_ALIGN_RIGHT:
				return 3;
			case BIN_ALIGN_LEFT:
				return 8;
			case BIN_ALIGN_CENTER:
				return 1;
			case BIN_ALIGN_JUSTIFY:
				return EJC_BOTH;
			case BIN_ALIGN_DISTRIBUTED:
				return EJC_DISTRIBUTE;
			case BIN_ALIGN_THAI_DISTRIBUTED:
				return EJC_THAI_DISTRIBUTE;
			default:
				return null;
		}
	}

	function collectRunText(oRun, oBuilt)
	{
		if (!oRun || !oRun.Content)
			return;
		for (var i = 0; i < oRun.Content.length; i++)
		{
			var oItem = oRun.Content[i];
			if (!oItem)
				continue;
			var nType = oItem.Type;
			if (nType === para_Text)
			{
				var nCode = oItem.GetCharCode ? oItem.GetCharCode() : oItem.Value;
				if (nCode == null)
					continue;
				oBuilt.owners.push(oItem);
				oBuilt.text += String.fromCharCode(nCode);
				if (isThaiScript(nCode))
					oBuilt.hasThai = true;
			}
			else if (nType === para_Space)
			{
				oBuilt.owners.push(null);
				oBuilt.text += " ";
			}
			else if (nType === para_Tab || nType === para_NewLine || nType === para_End)
			{
				oBuilt.owners.push(null);
				oBuilt.text += "\n";
			}
		}
	}

	function walkRuns(oContainer, fVisit)
	{
		var arr = oContainer && oContainer.Content;
		if (!arr || !arr.length)
			return;
		for (var i = 0; i < arr.length; i++)
		{
			var oItem = arr[i];
			if (!oItem)
				continue;
			if (oItem.Type === para_Run)
				fVisit(oItem);
			else if (oItem.Content)
				walkRuns(oItem, fVisit);
		}
	}

	function applyThaiLineBreaks(oParagraph)
	{
		if (!oParagraph || !oParagraph.Content)
			return;
		if (typeof para_Run === "undefined" || typeof para_Text === "undefined")
			return;

		var oBuilt = { text: "", owners: [], hasThai: false };
		walkRuns(oParagraph, function (oRun)
		{
			collectRunText(oRun, oBuilt);
		});

		var arrOwners = oParagraph._thaiBreakOwners;
		var bSame = oBuilt.text === oParagraph._thaiBreakText && arrOwners && arrOwners.length === oBuilt.owners.length;
		if (bSame)
		{
			for (var nOwn = 0; nOwn < arrOwners.length; nOwn++)
			{
				if (arrOwners[nOwn] !== oBuilt.owners[nOwn])
				{
					bSame = false;
					break;
				}
			}
		}
		if (bSame)
			return;

		var arrOld = oParagraph._thaiBreakItems;
		if (arrOld)
		{
			for (var nOld = 0; nOld < arrOld.length; nOld++)
			{
				var oOld = arrOld[nOld];
				if (oOld && oOld.SetSpaceAfter && oOld.private_IsSpaceAfter)
					oOld.SetSpaceAfter(!!oOld.private_IsSpaceAfter());
			}
		}

		oParagraph._thaiBreakText = oBuilt.text;
		oParagraph._thaiBreakOwners = oBuilt.owners;
		oParagraph._thaiBreakItems = [];
		if (!oBuilt.hasThai)
			return;

		var arrBreaks = thaiBreakAfterIndices(oBuilt.text);
		var oBreakAt = {};
		for (var nBreak = 0; nBreak < arrBreaks.length; nBreak++)
			oBreakAt[arrBreaks[nBreak]] = true;

		var arrMarked = [];
		var sText = oBuilt.text;
		for (var i = 0; i < oBuilt.owners.length; i++)
		{
			var oItem = oBuilt.owners[i];
			if (!oItem || !oBreakAt[i])
				continue;
			var nCode = sText.charCodeAt(i);
			if (!isThaiScript(nCode))
				continue;
			var nNext = i + 1 < sText.length ? sText.charCodeAt(i + 1) : 0;
			if (isThaiCombiningMark(nNext))
				continue;
			oItem.SetSpaceAfter(true);
			arrMarked.push(oItem);
		}
		oParagraph._thaiBreakItems = arrMarked;
	}

	AscWord.isThaiCombiningMark = isThaiCombiningMark;
	AscWord.isThaiScript = isThaiScript;
	AscWord.thaiBreakAfterIndices = thaiBreakAfterIndices;
	AscWord.dictionaryBreakAfterIndices = dictionaryBreakAfterIndices;
	AscWord.calcDistributeGaps = calcDistributeGaps;
	AscWord.clusterTakesJustify = clusterTakesJustify;
	AscWord.countGraphemeClusters = countGraphemeClusters;
	AscWord.alignFromWordJc = alignFromWordJc;
	AscWord.alignToWordJc = alignToWordJc;
	AscWord.alignFromDrawingAlgn = alignFromDrawingAlgn;
	AscWord.alignToDrawingAlgn = alignToDrawingAlgn;
	AscWord.paraAlignFromBinaryByte = paraAlignFromBinaryByte;
	AscWord.paraAlignToBinaryByte = paraAlignToBinaryByte;
	AscWord.x2tJcToByte = x2tJcToByte;
	AscWord.x2tByteToJc = x2tByteToJc;
	AscWord.x2tSetValueFromByte = x2tSetValueFromByte;
	AscWord.applyThaiLineBreaks = applyThaiLineBreaks;
	AscWord.BIN_ALIGN_JUSTIFY = BIN_ALIGN_JUSTIFY;
	AscWord.BIN_ALIGN_DISTRIBUTED = BIN_ALIGN_DISTRIBUTED;
	AscWord.BIN_ALIGN_THAI_DISTRIBUTED = BIN_ALIGN_THAI_DISTRIBUTED;

	if (typeof module === "object" && module.exports)
		module.exports = AscWord;
})(typeof window !== "undefined" ? window : this);

package hoboexcel

// BSD 3-Clause License

// Copyright (c) 2016 - 2018 360 Enterprise Security Group, Endpoint Security,
// inc. All rights reserved.

// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are met:

// * Redistributions of source code must retain the above copyright notice, this
//   list of conditions and the following disclaimer.

// * Redistributions in binary form must reproduce the above copyright notice,
//   this list of conditions and the following disclaimer in the documentation
//   and/or other materials provided with the distribution.

// * Neither the name of Excelize nor the names of its
//   contributors may be used to endorse or promote products derived from
//   this software without specific prior written permission.

// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
// AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
// DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
// FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
// DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
// SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
// CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
// OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
// OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

import "encoding/xml"

// xlsxSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main. String values may
// be stored directly inside spreadsheet cell elements; however, storing the
// same value inside multiple cell elements can result in very large worksheet
// Parts, possibly resulting in performance degradation. The Shared String Table
// is an indexed list of string values, shared across the workbook, which allows
// implementations to store values only once.
type xlsxSST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	SI          []xlsxSI `xml:"si"`
}

// xlsxSI directly maps the si element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - currently I have
// not checked this for completeness - it does as much as I need.
type xlsxSI struct {
	T string  `xml:"t"`
	R []xlsxR `xml:"r"`
}

// xlsxR directly maps the r element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - currently I have
// not checked this for completeness - it does as much as I need.
type xlsxR struct {
	//RPr *xlsxRPr `xml:"rPr"`
	T string `xml:"t"`
}

window.x_spreadsheet_file = (function() {
	var cssPrefix = 'x-spreadsheet';
	
	// ================================================================================
	// https://github.com/SheetJS/sheetjs/tree/master/demos/xspreadsheet
	function stox(wb) {
		var out = [];
		wb.SheetNames.forEach(function(name) {
			var o = {name:name, rows:{}};
			var ws = wb.Sheets[name];
			var aoa = XLSX.utils.sheet_to_json(ws, {raw: false, header:1});
			var maxRows = 1, maxCols = 1;
			aoa.forEach(function(r, i) {
				var cells = {};
				r.forEach(function(c, j) {
					cells[j] = ({ text: c });
					if (j >= maxCols) maxCols = j+1;
				});
				o.rows[i] = { cells: cells };
				if (i >= maxRows) maxRows = i+1;
			})
			o.rows.len = maxRows;
			o.cols = {len: maxCols};
			out.push(o);
		});
		return out;
	}
	
	function xwstoaoa(xws) {
		var aoa = [[]];
		var rowobj = xws.rows;
		for(var ri = 0; ri < rowobj.len; ++ri) {
			var row = rowobj[ri];
			if(!row) continue;
			aoa[ri] = [];
			Object.keys(row.cells).forEach(function(k) {
				var idx = +k;
				if(isNaN(idx)) return;
				aoa[ri][idx] = row.cells[k].text;
			});
		}
		return aoa;
	}
	// ================================================================================
	
	function getCurrentSheetIndex(xs) {
		var currentIndex = -1;
		xs.bottombar.items.some((item, i) => {
			if (item.el.classList.contains('active')) {
				currentIndex = i;
				return true;
			}
			return false;
		});
		return currentIndex;
	}
	
	function saveAsXlsx(xs) {
		var wb = XLSX.utils.book_new();
		xs.getData().forEach(function(xws) {
			var aoa = xwstoaoa(xws);
			var ws = XLSX.utils.aoa_to_sheet(aoa);
			XLSX.utils.book_append_sheet(wb, ws, xws.name);
		});
		XLSX.writeFile(wb, 'sheet.xlsx', {});
	}
	
	function saveAsCsv(xs) {
		alert('Only the active sheet is saved when saving as csv.');
		
		var ci = getCurrentSheetIndex(xs);
		var xws = xs.datas[ci].getData();
		var ws = XLSX.utils.aoa_to_sheet(xwstoaoa(xws));
		
		var wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, 'sheet');
		XLSX.writeFile(wb, xws.name+'.csv', {bookType:'csv'});
	}
	
	// ================================================================================
	return (function(xs) {
		// Add hidden input[type=file] tag to upload files
		var inputFile = document.createElement('input');
		inputFile.type = 'file';
		if(inputFile.addEventListener) {
			function handleFile(e) {
				var f = e.target.files[0];
				var reader = new FileReader();
				reader.onload = function(e) {
					var data = new Uint8Array(e.target.result);
					var xs_data = stox(XLSX.read(data, {type: 'array'}));
					xs.loadData(xs_data);
				};
				reader.readAsArrayBuffer(f);
				e.target.value = '';
			}
			inputFile.addEventListener('change', handleFile, false);
		}
		
		// Add "File" button in toolbar (REF: x-spreadsheet/src/component/toolbar.js)
		const toolbar = xs.sheet.toolbar;
		var Dropdown = xs.sheet.toolbar.moreEl.dd.__proto__.__proto__.constructor;
		var Element = Dropdown.__proto__;
		
		class DropdownFile extends Dropdown {
			constructor() {
				const children = [
					new Element('div', `${cssPrefix}-item`)
					.on('click', () => {
						inputFile.click();
						this.hide();
					})
					.child('Open...'),
					new Element('div', `${cssPrefix}-item`)
					.on('click', () => {
						this.hide();
						saveAsXlsx(xs);
					})
					.child('Save (Ctrl+S)'),
					new Element('div', `${cssPrefix}-item`)
					.on('click', () => {
						this.hide();
						saveAsCsv(xs);
					})
					.child('Save as CSV (Ctrl+Shift+S)')
				];
				super('File', '200px', true, 'bottom-left', ...children);
			}
		}

		toolbar.ddFile = new DropdownFile();
		const newChildren = [
			new Element('div', `${cssPrefix}-toolbar-divider`),
			new Element('div', `${cssPrefix}-toolbar-btn`).child(toolbar.ddFile)
		];
		setTimeout(() => {
			newChildren.forEach((it) => {
				toolbar.btns.el.prepend(it.el);
				const rect = it.box();
				const { marginLeft, marginRight } = it.computedStyle();
				toolbar.btns2.unshift([it, rect.width + parseInt(marginLeft, 10) + parseInt(marginRight, 10)]);
			});
		}, 0);
		
		// Bind keydown event (REF: x-spreadsheet/src/component/sheet.js)
		window.addEventListener('keydown', (evt) => {
			const keyCode = evt.keyCode || evt.which;
			if ((evt.ctrlKey || evt.metaKey) && keyCode == 83) { // Ctrl + S
				if (!evt.shiftKey) {
					saveAsXlsx(xs);
				}else{
					saveAsCsv(xs);
				}
				evt.preventDefault();
			}
		})
	})
}());

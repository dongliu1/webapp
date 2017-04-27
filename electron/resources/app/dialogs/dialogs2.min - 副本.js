var __extends = this.__extends ||
function(n, t) {
	function r() {
		this.constructor = n
	}
	for (var i in t) t.hasOwnProperty(i) && (n[i] = t[i]);
	r.prototype = t.prototype, n.prototype = new r
}, sku;
(function(n) {
	(function(t) {
		(function(i) {
			var ft, o, et, ot, ut, tt, it, bt, pt, rt, st, vt, yt, ct, ht, at, k, lt, l, c, v, a, e, u, h, s, d, f, nt, g, p, y, b, w, r;
			var wt = function(n) {
					function r() {
						var t = this,
							r = {
								width: "auto",
								title: i.res.seriesDialog.series,
								modal: !0,
								buttons: [{
									text: i.res.ok,
									click: function() {
										t.fill(), t.close(), i.wrapper.setFocusToSpread()
									}
								}, {
									text: i.res.cancel,
									click: function() {
										t.close(), i.wrapper.setFocusToSpread()
									}
								}]
							};
						n.call(this, "./dialogs/dialogs2.html", ".fill-dialog", r)
					}
					__extends(r, n);
					r.prototype._init = function() {
						this._seriesColumns = this._element.find("#series-columns"), this._seriesRows = this._element.find("#series-rows"), this._typeLinear = this._element.find("#type-linear"), this._typeGrowth = this._element.find("#type-growth"), this._typeDate = this._element.find("#type-date"), this._typeAutofill = this._element.find("#type-autofill"), this._dateDay = this._element.find("#date-day"), this._dateWeekDay = this._element.find("#date-weekday"), this._dateMonth = this._element.find("#date-month"), this._dateYear = this._element.find("#date-year"), this._seriesTrend = this._element.find("#series-trend"), this._stepValue = this._element.find("#step-value"), this._stopValue = this._element.find("#stop-value"), this._fillSeries = 0, this._fillType = 1, this._fillDateUnit = 0, this._seriesColumns.prop("checked", !0), this._typeLinear.prop("checked", !0), this._dateDay.prop("checked", !0), this._stepValue.val(1), this._disableDate();
						var n = this;
						this._seriesColumns.on("click", function() {
							n._fillSeries = 0
						});
						this._seriesRows.on("click", function() {
							n._fillSeries = 1
						});
						this._typeLinear.on("click", function() {
							n._seriesTrend.prop("checked") || (n._disableDate(), n._enableStepStop()), n._fillType = 1
						});
						this._typeGrowth.on("click", function() {
							n._seriesTrend.prop("checked") || (n._disableDate(), n._enableStepStop()), n._fillType = 2
						});
						this._typeAutofill.on("click", function() {
							n._disableDate(), n._disableStepStop(), n._fillType = 4
						});
						this._typeDate.on("click", function() {
							n._enableDate(), n._enableDate(), n._fillType = 3
						});
						this._seriesTrend.on("change", function() {
							$(this).prop("checked") ? n._enableTrend() : n._disableTrend()
						});
						this._dateDay.on("click", function() {
							n._fillDateUnit = 0
						});
						this._dateWeekDay.on("click", function() {
							n._fillDateUnit = 1
						});
						this._dateMonth.on("click", function() {
							n._fillDateUnit = 2
						});
						this._dateYear.on("click", function() {
							n._fillDateUnit = 3
						})
					}					r.prototype._disableDate = function() {
						this._dateDay.attr("disabled", "disabled"), this._dateWeekDay.attr("disabled", "disabled"), this._dateMonth.attr("disabled", "disabled"), this._dateYear.attr("disabled", "disabled")
					}					r.prototype._enableDate = function() {
						this._dateDay.removeAttr("disabled"), this._dateWeekDay.removeAttr("disabled"), this._dateMonth.removeAttr("disabled"), this._dateYear.removeAttr("disabled")
					}					r.prototype._disableStepStop = function() {
						this._stepValue.attr("disabled", "disabled"), this._stopValue.attr("disabled", "disabled")
					}					r.prototype._enableStepStop = function() {
						this._stepValue.removeAttr("disabled"), this._stopValue.removeAttr("disabled")
					}					r.prototype._disableTrend = function() {
						this._typeDate.removeAttr("disabled"), this._typeAutofill.removeAttr("disabled"), this._enableStepStop()
					}					r.prototype._enableTrend = function() {
						this._typeLinear.prop("checked", !0), this._typeLinear.removeAttr("disabled"), this._typeGrowth.removeAttr("disabled"), this._typeDate.attr("disabled", "disabled"), this._typeAutofill.attr("disabled", "disabled"), this._disableDate(), this._disableStepStop()
					}					r.prototype.fill = function() {
						var r = i.wrapper.spread.getActiveSheet(),
							s, o, u, n, e, f;
						for (r.isPaintSuspended(!0), r.suspendCalcService(), s = r.getSelections(), o = 0; o < s.length; o++) {
							n = s[o], u = this._fillSeries == 0 ? new t.Range(n.row, n.col, 1, n.colCount) : new t.Range(n.row, n.col, n.rowCount, 1);
							if (this._seriesTrend.prop("checked")) switch (this._fillType) {
							case 1:
								r.fillLinear(u, n, this._fillSeries, undefined, undefined);
								break;
							case 2:
								r.fillGrowth(u, n, this._fillSeries, undefined, undefined);
								break
							} else {
								e = Number(this._stepValue.val()), f = this._stopValue.val();
								switch (this._fillType) {
								case 1:
									f ? (f = Number(this._stopValue.val()), r.fillLinear(u, n, this._fillSeries, e, f)) : r.fillLinear(u, n, this._fillSeries, e, undefined);
									break;
								case 2:
									f ? (f = Number(this._stopValue.val()), r.fillGrowth(u, n, this._fillSeries, e, f)) : r.fillGrowth(u, n, this._fillSeries, e, undefined);
									break;
								case 4:
									r.fillAuto(u, n, this._fillSeries);
									break;
								case 3:
									f ? (f = new Date(this._stopValue.val()), r.fillDate(u, n, this._fillSeries, this._fillDateUnit, e, f)) : r.fillDate(u, n, this._fillSeries, this._fillDateUnit, e);
									break
								}
							}
						}
						r.isPaintSuspended(!1), r.resumeCalcService()
					} 
					return r;
				}(i.BaseDialog); i.FillDialog = wt;
			ft = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.customSortDialog.sort,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.sort(), t.close(), i.wrapper.setFocusToSpread()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".sort-dialog", r)
				}
				__extends(t, n);
				t.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet();
					this._sortByOptions = this._getSortByOptions(), this._element.find(".sort-infolist").empty(), this._addLevel()
				}
				t.prototype._getSortByOptions = function() {
					var f = i.wrapper.spread.getActiveSheet(),
						o = f._selectionModel[f._selectionModel.activeSelectedRangeIndex],
						e = $("<select><select>"),
						u, r, n, t;
					e.addClass("sort-row-sortby");
					if (this._byRows) for (u = o.col, u = u == -1 ? 0 : u, n = 0; n < o.colCount; n++) t = $("<option></option>"), t.text("Column " + f.getCell(f.getRowCount(1) - 1, u + n, 1).text()), t.val(u + n), e.append(t);
					else for (r = o.row, r = r == -1 ? 0 : r, n = 0; n < o.rowCount; n++) t = $("<option></option>"), t.text("Row " + (r + n + 1).toString()), t.val(r + n), e.append(t);
					return e
				}				t.prototype._init = function() {
					var n = this;
					this._byRows = !0, this._ascending = !0;
					this._element.find("#add-level").on("click", function() {
						n._addLevel()
					});
					this._element.find("#delete-level").on("click", function() {
						n._deleteLevel()
					});
					this._element.find("#copy-level").on("click", function() {
						n._copyLevel()
					});
					this._element.find("#move-up").on("click", function() {
						n._moveUp()
					});
					this._element.find("#move-down").on("click", function() {
						n._moveDown()
					});
					this._element.find("#sort-options").on("click", function() {
						n._sortOption()
					})
				}				t.prototype._addLevel = function() {
					var n = $("<Div></Div>").addClass("sort-info").appendTo(this._element.find(".sort-infolist")),
						t;
					n.parent().children().length == 1 ? (this._activeSortInfo = n, this._activeSortInfo.addClass("ui-state-default"), n.append($("<label></label>").addClass("sort-row-header").text(i.res.customSortDialog.sortBy2))) : n.append($("<label></label>").addClass("sort-row-header").text(i.res.customSortDialog.thenBy)), n.append(this._sortByOptions.clone(!0)), n.append($("<select></select>").addClass("sort-column").append($("<option>Values</option>"))), n.append($("<select></select>").addClass("sort-column sort-order").append($("<option>A to Z</option><option>Z to A</option>"))), n.append($("<div></div>").addClass("clear-float")), t = this, n.click(function() {
						t._activeSortInfo.removeClass("ui-state-default"), t._activeSortInfo = $(this), t._activeSortInfo.addClass("ui-state-default")
					})
				}				t.prototype._deleteLevel = function() {
					var n = this._activeSortInfo;
					this._activeSortInfo = n.prev(), this._activeSortInfo.length == 0 && (this._activeSortInfo = n.next(), this._activeSortInfo.find(".sort-row-header").text(i.res.customSortDialog.sortBy2)), this._activeSortInfo.addClass("ui-state-default"), n.remove()
				}				t.prototype._copyLevel = function() {
					var n = this._activeSortInfo.clone(!0);
					this._activeSortInfo.prev().length == 0 && n.find(".sort-row-header").text(i.res.customSortDialog.thenBy), n.find(".sort-row-sortby").val(this._activeSortInfo.find(".sort-row-sortby").val()), n.find(".sort-order").val(this._activeSortInfo.find(".sort-order").val()), this._activeSortInfo.after(n), this._activeSortInfo.removeClass("ui-state-default"), this._activeSortInfo = this._activeSortInfo.next()
				}				t.prototype._moveUp = function() {
					this._activeSortInfo.prev().before(this._activeSortInfo), this._updateSortHeader()
				}				t.prototype._moveDown = function() {
					this._activeSortInfo.next().after(this._activeSortInfo), this._updateSortHeader()
				}				t.prototype._sortOption = function() {
					var n = this;
					this._sortOptionDialog || (this._sortOptionDialog = new o(n)), this._sortOptionDialog.open()
				}				t.prototype._setByRows = function(n) {
					if (this._byRows == n) return;
					else this._byRows = n, this._sortByOptions = this._getSortByOptions(), this._element.find(".sort-infolist").empty(), this._addLevel()
				}				t.prototype._updateSortHeader = function() {
					this._element.find(".sort-infolist").children().each(function(n, t) {
						n == 0 ? $(t).find(".sort-row-header").text(i.res.customSortDialog.sortBy2) : $(t).find(".sort-row-header").text(i.res.customSortDialog.thenBy)
					})
				}				t.prototype.sort = function() {
					var t = i.wrapper.spread.getActiveSheet(),
						n = t._selectionModel[t._selectionModel.activeSelectedRangeIndex],
						u = this._element.find(".sort-infolist"),
						r = [];
					if (u.length == 0) return;
					u.children().each(function(n, t) {
						var f = !0,
							u, i;
						$(t).find(".sort-order").val() == "Z to A" && (f = !1), u = $(t).find(".sort-row-sortby").val(), i = {
							index: u,
							ascending: f
						}, r.push(i)
					}), t.sortRange(n.row, n.col, n.rowCount, n.colCount, this._byRows, r)
				}
				return t;
			}(i.BaseDialog);i.SortDialog = ft;
			
			SortOptionDialog = function(n) {
				function t(t) {
					var r = this,
						u = {
							width: "auto",
							modal: !0,
							title: i.res.customSortDialog.sortOptions,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._setByRows(r._byRows), r.close()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									r.close()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".sort-option-dialog", u)
				}
				__extends(t, n);
				t.prototype._init = function() {
					var n = this;
					this._byRows = !0, this._element.find("#sort-byrows").prop("checked", !0).click(function() {
						n._byRows = !0
					}), this._element.find("#sort-bycols").click(function() {
						n._byRows = !1
					})
				}
				return t;
			}(i.BaseDialog); i.SortOptionDialog = SortOptionDialog;
			et = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							resizable: !1,
							title: i.res.spreadSettingDialog.spreadSetting,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting(), i.actions.isFileModified = !0, t.close(), i.wrapper.setFocusToSpread(), i.ribbon.updateShowHide()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".spread-setting-dialog", r)
				}
				return __extends(t, n), t.prototype._init = function() {
					this.settingTab = this._element.find(".spread-setting-tab").tabs()
				}, t.prototype._beforeOpen = function(n) {
					var t, r;
					switch (n[0]) {
					case "general":
						t = 0;
						break;
					case "scrollbar":
						t = 1;
						break;
					case "tabstrip":
						t = 2;
						break
					}
					this.settingTab.tabs("option", "active", t), this._element.find("#setting-dropdrop").prop("checked", i.wrapper.spread.canUserDragDrop()), this._element.find("#setting-dragfill").prop("checked", i.wrapper.spread.canUserDragFill()), this._element.find("#setting-undo").prop("checked", i.wrapper.spread.allowUndo()), this._element.find("#setting-formula").prop("checked", i.wrapper.spread.canUserEditFormula()), this._element.find("#setting-zoom").prop("checked", i.wrapper.spread.allowUserZoom()), this._element.find("#setting-overflow").prop("checked", i.wrapper.spread.getActiveSheet().allowCellOverflow()), this._element.find("#setting-vertical-scrollbar").prop("checked", i.wrapper.spread.showVerticalScrollbar()), this._element.find("#setting-horizontal-scrollbar").prop("checked", i.wrapper.spread.showHorizontalScrollbar()), this._element.find("#setting-scrollbar-showmax").prop("checked", i.wrapper.spread.scrollbarShowMax()), this._element.find("#setting-scrollbar-maxalign").prop("checked", i.wrapper.spread.scrollbarMaxAlign()), this._element.find("#setting-tab-visible").prop("checked", i.wrapper.spread.tabStripVisible()), this._element.find("#setting-new-tab").prop("checked", i.wrapper.spread.newTabVisible()), this._element.find("#setting-tab-edit").prop("checked", i.wrapper.spread.tabEditable()), r = this._element.find("#setting-tab-ratio"), r.spinner(), r.val(Math.round(i.wrapper.spread.getTabStripRatio() * 100))
				}, t.prototype._applySetting = function() {
					var n, t;
					for (i.wrapper.spread.canUserDragDrop(this._element.find("#setting-dropdrop").prop("checked")), i.wrapper.spread.canUserDragFill(this._element.find("#setting-dragfill").prop("checked")), i.wrapper.spread.allowUndo(this._element.find("#setting-undo").prop("checked")), i.wrapper.spread.canUserEditFormula(this._element.find("#setting-formula").prop("checked")), i.wrapper.spread.allowUserZoom(this._element.find("#setting-zoom").prop("checked")), n = 0; n < i.wrapper.spread.sheets.length; n++) i.wrapper.spread.sheets[n].allowCellOverflow(this._element.find("#setting-overflow").prop("checked"));
					i.wrapper.spread.showVerticalScrollbar(this._element.find("#setting-vertical-scrollbar").prop("checked")), i.wrapper.spread.showHorizontalScrollbar(this._element.find("#setting-horizontal-scrollbar").prop("checked")), i.wrapper.spread.scrollbarShowMax(this._element.find("#setting-scrollbar-showmax").prop("checked")), i.wrapper.spread.scrollbarMaxAlign(this._element.find("#setting-scrollbar-maxalign").prop("checked")), i.wrapper.spread.tabStripVisible(this._element.find("#setting-tab-visible").prop("checked")), i.wrapper.spread.newTabVisible(this._element.find("#setting-new-tab").prop("checked")), i.wrapper.spread.tabEditable(this._element.find("#setting-tab-edit").prop("checked")), t = this._element.find("#setting-tab-ratio").val(), $.isNumeric(t) && i.wrapper.spread.setTabStripRatio(t / 100, !0)
				}, t
			}(i.BaseDialog), i.SpreadSettingDialog = et, ot = function(n) {
				function r() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							resizable: !1,
							title: i.res.sheetSettingDialog.sheetSetting,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.close(), t._applySetting(), i.actions.isFileModified = !0, i.wrapper.setFocusToSpread(), i.ribbon.updateShowHide()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".sheet-setting-dialog", r)
				}
				__extends(r, n);
				r.prototype._init = function() {
					var n = this;
					this._element.find(".sheet-setting-tab").tabs(), this._element.find(".header-setting-tab").tabs(), this._element.find("#sheet-column-count").spinner(), this._element.find("#sheet-row-count").spinner(), this._element.find("#frozen-column-count").spinner(), this._element.find("#frozen-row-count").spinner(), this._element.find("#trailing-column-count").spinner(), this._element.find("#trailing-row-count").spinner(), this._element.find("#header-row-count").spinner(), this._element.find("#header-row-autoindex").spinner(), this._element.find("#header-row-default-width").spinner(), this._element.find("#header-column-count").spinner(), this._element.find("#header-column-autoindex").spinner(), this._element.find("#header-column-default-height").spinner(), this._element.find("#gridline-color-picker").wijcolorpicker({
						valueChanged: function(t, i) {
							i.color === undefined ? n._element.find("#gridline-color-span").css("background-color", "") : n._element.find("#gridline-color-span").css("background-color", i.color)
						},
						choosedColor: function(t, i) {
							n._element.find("#gridline-color-frame").wijcomboframe("close")
						},
						openColorDialog: function(t, i) {
							n._element.find("#gridline-color-frame").wijcomboframe("close")
						},
						showNoFill: !1
					}), this._element.find("#gridline-color-frame").wijcomboframe(), this._element.find("#sheettab-color-picker").wijcolorpicker({
						valueChanged: function(t, i) {
							i.color === undefined ? n._element.find("#sheettab-color-span").css("background-color", "") : n._element.find("#sheettab-color-span").css("background-color", i.color)
						},
						choosedColor: function(t, i) {
							n._element.find("#sheettab-color-frame").wijcomboframe("close")
						},
						openColorDialog: function(t, i) {
							n._element.find("#sheettab-color-frame").wijcomboframe("close")
						},
						showNoFill: !1
					}), this._element.find("#sheettab-color-frame").wijcomboframe();
					this._element.find("#reference-A1").on("click", function() {
						n._element.find("#header-column-autotext-letters").prop("checked", !0)
					});
					this._element.find("#reference-R1C1").on("click", function() {
						n._element.find("#header-column-autotext-numbers").prop("checked", !0)
					})
				}				r.prototype._beforeOpen = function(n) {
					var e = this,
						r = i.wrapper.spread.getActiveSheet(),
						f, u;
					$("#gridline-color-picker").wijcolorpicker({
						themeColors: i.wrapper.getThemeColors()
					}), $("#sheettab-color-picker").wijcolorpicker({
						themeColors: i.wrapper.getThemeColors()
					});
					switch (n[0]) {
					case "general":
						f = 0;
						break;
					case "gridline":
						f = 1;
						break;
					case "calculation":
						f = 2;
						break;
					case "headers":
						f = 3;
						break;
					case "sheettab":
						f = 4;
						break
					}
					this._element.find(".sheet-setting-tab").tabs("option", "active", f), this._element.find("#sheet-column-count").val(r.getColumnCount()), this._element.find("#sheet-row-count").val(r.getRowCount()), this._element.find("#frozen-column-count").val(r.getFrozenColumnCount()), this._element.find("#frozen-row-count").val(r.getFrozenRowCount()), this._element.find("#trailing-column-count").val(r.getFrozenTrailingColumnCount()), this._element.find("#trailing-row-count").val(r.getFrozenTrailingRowCount()), this._element.find("input[name='selection-policy']").each(function(n, i) {
						$(i).val() == t.SelectionPolicy[r.selectionPolicy()] && $(i).prop("checked", !0)
					}), this._element.find("#sheet-protect").prop("checked", r.getIsProtected()), this._element.find("#gridline-horizontal").prop("checked", r.getGridlineOptions().showHorizontalGridline), this._element.find("#gridline-vertical").prop("checked", r.getGridlineOptions().showVerticalGridline), r.getGridlineOptions().color ? (u = i.ColorHelper.parse(r.getGridlineOptions().color, r.currentTheme().colors()), $("#gridline-color-picker").wijcolorpicker("option", "selectedItem", u), this._element.find("#gridline-color-span").css("background-color", u.color)) : (this._element.find("#gridline-color-span").css("background-color", "transparent"), $("#gridline-color-picker").wijcolorpicker("option", "selectedItem", null)), this._element.find("input[name='reference-style']").each(function(n, i) {
						$(i).val() == t.ReferenceStyle[r.referenceStyle()] && $(i).prop("checked", !0)
					}), this._element.find("#header-column-count").val(r.getRowCount(1)), this._element.find("input[name='column-autotext']").each(function(n, i) {
						$(i).val() == t.HeaderAutoText[r.getColumnHeaderAutoText()] && $(i).prop("checked", !0)
					}), this._element.find("#header-column-autoindex").val(r.getColumnHeaderAutoTextIndex()), this._element.find("#header-column-default-height").val(r.defaults.colHeaderRowHeight), this._element.find("#header-column-visible").prop("checked", r.getColumnHeaderVisible()), this._element.find("#header-row-count").val(r.getColumnCount(2)), this._element.find("input[name='row-autotext']").each(function(n, i) {
						$(i).val() == t.HeaderAutoText[r.getRowHeaderAutoText()] && $(i).prop("checked", !0)
					}), this._element.find("#header-row-autoindex").val(r.getRowHeaderAutoTextIndex()), this._element.find("#header-row-default-width").val(r.defaults.rowHeaderColWidth), this._element.find("#header-row-visible").prop("checked", r.getRowHeaderVisible()), r.sheetTabColor() ? (u = i.ColorHelper.parse(r.sheetTabColor(), r.currentTheme().colors()), $("#sheettab-color-picker").wijcolorpicker("option", "selectedItem", u), this._element.find("#sheettab-color-span").css("background-color", u.color)) : (this._element.find("#sheettab-color-span").css("background-color", "transparent"), $("#sheettab-color-picker").wijcolorpicker("option", "selectedItem", null))
				}				r.prototype._applySetting = function() {
					var n = i.wrapper.spread.getActiveSheet(),
						u, r;
					n.isPaintSuspended(!0), n.suspendCalcService(), n.setColumnCount(this._element.find("#sheet-column-count").val()), n.setRowCount(this._element.find("#sheet-row-count").val()), n.setFrozenRowCount(Number(this._element.find("#frozen-row-count").val())), n.setFrozenColumnCount(Number(this._element.find("#frozen-column-count").val())), n.setFrozenTrailingRowCount(Number(this._element.find("#trailing-row-count").val())), n.setFrozenTrailingColumnCount(Number(this._element.find("#trailing-column-count").val())), n.selectionPolicy(t.SelectionPolicy[this._element.find("input[name='selection-policy']:checked").val()]), n.setIsProtected(this._element.find("#sheet-protect").prop("checked")), u = $("#gridline-color-picker").wijcolorpicker("option", "selectedItem"), u && n.setGridlineOptions({
						showHorizontalGridline: this._element.find("#gridline-horizontal").prop("checked"),
						showVerticalGridline: this._element.find("#gridline-vertical").prop("checked"),
						color: u.color
					}), n.referenceStyle(t.ReferenceStyle[this._element.find("input[name='reference-style']:checked").val()]), n.setRowCount(this._element.find("#header-column-count").val(), 1), n.setColumnHeaderAutoText(t.HeaderAutoText[this._element.find("input[name='column-autotext']:checked").val()]), n.setColumnHeaderAutoTextIndex(Number(this._element.find("#header-column-autoindex").val())), n.defaults.colHeaderRowHeight = this._element.find("#header-column-default-height").val(), n.setColumnHeaderVisible(this._element.find("#header-column-visible").prop("checked")), n.setColumnCount(this._element.find("#header-row-count").val(), 2), n.setRowHeaderAutoText(t.HeaderAutoText[this._element.find("input[name='row-autotext']:checked").val()]), n.setRowHeaderAutoTextIndex(Number(this._element.find("#header-row-autoindex").val())), n.defaults.rowHeaderColWidth = this._element.find("#header-row-default-width").val(), n.setRowHeaderVisible(this._element.find("#header-row-visible").prop("checked")), r = $("#sheettab-color-picker").wijcolorpicker("option", "selectedItem"), r && (r.name ? n.sheetTabColor(r.name) : n.sheetTabColor(r.color)), n.isPaintSuspended(!1), n.resumeCalcService()
				}
				return r;
			}(i.BaseDialog), i.SheetSettingDialog = ot, ut = function(r) {
				function u() {
					var n = this,
						t = {
							width: "auto",
							modal: !0,
							title: i.res.dataValidationDialog.dataValidation,
							buttons: [{
								text: i.res.dataValidationDialog.clearAll,
								click: function() {
									n._clearAll()
								}
							}, {
								text: i.res.ok,
								click: function() {
									n._ValidValues() && (n._applySetting(), n.close(), i.wrapper.setFocusToSpread())
								}
							}, {
								text: i.res.cancel,
								click: function() {
									n.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					r.call(this, "./dialogs/dialogs2.html", ".data-validation-dialog", t)
				}
				return __extends(u, r), u.prototype._init = function() {
					var r = this;
					this._element.find(".data-validation-tab").tabs(), this._criteriaType = this._element.find(".validation-criteria-type"), this._ignoreBlank = this._element.find(".validation-ignore-blank").prop("disabled", !0).prop("checked", !0), this._comparisonOperator = this._element.find(".validation-comparison-operator").prop("disabled", !0), this._inCellDropDown = this._element.find(".validation-incell-dropdown").prop("checked", !0), this._value1 = this._element.find(".validation-value1"), this._value2 = this._element.find(".validation-value2"), this._value1Label = this._element.find(".validation-value1-label"), this._value2Label = this._element.find(".validation-value2-label"), this._showInputMessage = this._element.find(".show-input-message").prop("checked", !0), this._inputTitle = this._element.find(".input-title"), this._inputMessage = this._element.find(".input-message"), this._showErrorMessage = this._element.find(".show-error-message").prop("checked", !0), this._errorTitle = this._element.find(".error-title"), this._errorMessage = this._element.find(".error-message"), this._errorStyle = this._element.find(".error-style");
					this._criteriaType.on("change", function() {
						r._criteriaTypeChanged()
					});
					this._comparisonOperator.on("change", function() {
						r._comparisonOperatorChanged()
					});
					this._showInputMessage.on("change", function() {
						r._showInputMessageChanged()
					});
					this._showErrorMessage.on("change", function() {
						r._showErrorMessageChanged()
					});
					this._errorStyle.on("change", function() {
						r._errorStyleChanged()
					});
					i.wrapper.spread.bind(t.Events.ValidationError, n.data, function(n, t) {
						r._validationError(t)
					})
				}, u.prototype._beforeOpen = function() {
					this._element.find(".data-validation-tab").tabs("option", "active", 0);
					var t = i.wrapper.spread.getActiveSheet(),
						r = t._selectionModel[t._selectionModel.activeSelectedRangeIndex],
						n = t.getDataValidator(r.row, r.col);
					n ? (this._loadValidatorSetting(n), this._UpdateCriteriaStatus(n.type, n.comparisonOperator), this._updateErrorIcon()) : this._clearAll()
				}, u.prototype._loadValidatorSetting = function(n) {
					this._criteriaType.val(t.CriteriaType[n.type]), this._comparisonOperator.val(t.ComparisonOperator[n.comparisonOperator]), this._value1.val(n.value1()), this._value2.val(n.value2()), this._ignoreBlank.prop("checked", n.ignoreBlank), this._inCellDropDown.prop("checked", n.inCellDropdown), this._showInputMessage.prop("checked", n.showInputMessage), this._inputTitle.val(n.inputTitle), this._inputMessage.val(n.inputMessage), this._showErrorMessage.prop("checked", n.showErrorMessage), this._errorStyle.val(t.ErrorStyle[n.errorStyle]), this._errorTitle.val(n.errorTitle), this._errorMessage.val(n.errorMessage)
				}, u.prototype._clearAll = function() {
					this._resetCriteriaToDefaultStatus(), this._criteriaType.find("option:first-child").prop("selected", !0), this._comparisonOperator.find("option:first-child").prop("selected", !0), this._value1.val(null), this._value2.val(null), this._showInputMessage.prop("checked", !0), this._inputTitle.val(null).prop("disabled", !1), this._inputMessage.val(null).prop("disabled", !1), this._showErrorMessage.prop("checked", !0), this._errorStyle.prop("disabled", !1).find("option:first-child").prop("selected", !0), this._errorTitle.val(null).prop("disabled", !1), this._errorMessage.val(null).prop("disabled", !1), this._updateErrorIcon()
				}, u.prototype._resetCriteriaToDefaultStatus = function() {
					this._comparisonOperator.prop("disabled", !0), this._ignoreBlank.prop("disabled", !0), this._inCellDropDown.parent().css("visibility", "hidden"), this._value1.css("visibility", "hidden"), this._value1Label.css("visibility", "hidden"), this._value2.css("visibility", "hidden"), this._value2Label.css("visibility", "hidden")
				}, u.prototype._disableIgnoreBlank = function(n) {
					this._ignoreBlank.prop("disabled", n)
				}, u.prototype._disableComparisonOperator = function(n) {
					this._comparisonOperator.prop("disabled", n)
				}, u.prototype._showCellDropDown = function(n) {
					this._inCellDropDown.parent().css("visibility", n ? "visible" : "hidden")
				}, u.prototype._updateValue1 = function(n, t) {
					this._value1.css("visibility", n ? "visible" : "hidden"), this._value1Label.css("visibility", n ? "visible" : "hidden"), this._value1Label.text(t)
				}, u.prototype._updateValue2 = function(n, t) {
					this._value2.css("visibility", n ? "visible" : "hidden"), this._value2Label.css("visibility", n ? "visible" : "hidden"), this._value2Label.text(t)
				}, u.prototype._UpdateCriteriaStatus = function(n, t) {
					this._resetCriteriaToDefaultStatus();
					if (n == 0) return;
					else if (n == 1 || n == 2 || n == 4 || n == 6) {
						var u, r, f;
						switch (n) {
						case 1:
						case 2:
							u = i.res.dataValidationDialog.minimum, r = i.res.dataValidationDialog.maximum, f = i.res.dataValidationDialog.value;
							break;
						case 4:
							u = i.res.dataValidationDialog.startDate, r = i.res.dataValidationDialog.endDate, f = i.res.dataValidationDialog.dateLabel;
							break;
						case 6:
							u = i.res.dataValidationDialog.minimum, r = i.res.dataValidationDialog.maximum, f = i.res.dataValidationDialog.length;
							break
						}
						switch (t) {
						case 6:
						case 7:
							this._disableIgnoreBlank(!1), this._disableComparisonOperator(!1), this._showCellDropDown(!1), this._updateValue1(!0, u), this._updateValue2(!0, r);
							break;
						case 0:
						case 1:
							this._disableIgnoreBlank(!1), this._disableComparisonOperator(!1), this._showCellDropDown(!1), this._updateValue1(!0, f), this._updateValue2(!1, "");
							break;
						case 2:
						case 3:
							this._disableIgnoreBlank(!1), this._disableComparisonOperator(!1), this._showCellDropDown(!1), this._updateValue1(!0, u), this._updateValue2(!1, "");
							break;
						case 4:
						case 5:
							this._disableIgnoreBlank(!1), this._disableComparisonOperator(!1), this._showCellDropDown(!1), this._updateValue1(!0, r), this._updateValue2(!1, "");
							break
						}
					} else n == 3 ? (this._disableIgnoreBlank(!1), this._disableComparisonOperator(!0), this._showCellDropDown(!0), this._updateValue1(!0, i.res.dataValidationDialog.source), this._updateValue2(!1, "")) : n == 7 && (this._disableIgnoreBlank(!1), this._disableComparisonOperator(!0), this._showCellDropDown(!1), this._updateValue1(!0, i.res.dataValidationDialog.formula), this._updateValue2(!1, ""))
				}, u.prototype._criteriaTypeChanged = function() {
					this._UpdateCriteriaStatus(t.CriteriaType[this._criteriaType.val()], t.ComparisonOperator[this._comparisonOperator.val()])
				}, u.prototype._comparisonOperatorChanged = function() {
					this._UpdateCriteriaStatus(t.CriteriaType[this._criteriaType.val()], t.ComparisonOperator[this._comparisonOperator.val()])
				}, u.prototype._showInputMessageChanged = function() {
					this._inputTitle.prop("disabled", !this._showInputMessage.prop("checked")), this._inputMessage.prop("disabled", !this._showInputMessage.prop("checked"))
				}, u.prototype._showErrorMessageChanged = function() {
					this._errorTitle.prop("disabled", !this._showErrorMessage.prop("checked")), this._errorMessage.prop("disabled", !this._showErrorMessage.prop("checked")), this._errorStyle.prop("disabled", !this._showErrorMessage.prop("checked"))
				}, u.prototype._updateErrorIcon = function() {
					var n = this._element.find(".error-icon");
					switch (this._errorStyle.val()) {
					case "Stop":
						n.removeClass("error-icon-warning").removeClass("error-icon-info").addClass("error-icon-stop");
						break;
					case "Warning":
						n.removeClass("error-icon-stop").removeClass("error-icon-info").addClass("error-icon-warning");
						break;
					case "Information":
						n.removeClass("error-icon-warning").removeClass("error-icon-stop").addClass("error-icon-info");
						break
					}
				}, u.prototype._errorStyleChanged = function() {
					this._updateErrorIcon()
				}, u.prototype._validationError = function(n) {
					var u = n.sheet.getValue(n.row, n.col),
						r = n.validator,
						f = r.errorTitle,
						t = r.errorMessage;
					f || (f = i.res.title), r.showErrorMessage ? (n.validationResult = 0, r.errorStyle == 0 && (t || (t = i.res.dataValidationDialog.errorMessage1), i.MessageBox.show(t, f, 3, 1, function(t, i) {
						switch (i) {
						case 1:
							n.sheet.setActiveCell(n.row, n.col), n.sheet.setValue(n.row, n.col, u), n.sheet.startEdit(!0);
							break;
						case 4:
							n.sheet.setValue(n.row, n.col, u);
							break
						}
					})), r.errorStyle == 1 && (t || (t = i.res.dataValidationDialog.errorMessage2), i.MessageBox.show(t, f, 2, 2, function(t, i) {
						switch (i) {
						case 2:
							break;
						case 3:
							n.sheet.setActiveCell(n.row, n.col), n.sheet.setValue(n.row, n.col, u), n.sheet.startEdit(!0);
							break;
						case 4:
							n.sheet.setValue(n.row, n.col, u);
							break
						}
					})), r.errorStyle == 2 && (t || (t = i.res.dataValidationDialog.errorMessage1), i.MessageBox.show(t, f, 1, 1, function(t, i) {
						switch (i) {
						case 1:
							break;
						case 4:
							n.sheet.setValue(n.row, n.col, u);
							break
						}
					}))) : n.validationResult = 0
				}, u.prototype._applySetting = function() {
					var n, f = t.ComparisonOperator[this._comparisonOperator.val()],
						r = this._value1.val(),
						u = this._value2.val();
					switch (this._criteriaType.val()) {
					case "AnyValue":
						n = new t.DefaultDataValidator, n.type = 0;
						break;
					case "WholeNumber":
						$.isNumeric(r) && (r = Number(r)), $.isNumeric(u) && (u = Number(u)), n = t.DefaultDataValidator.createNumberValidator(f, r, u, !0);
						break;
					case "DecimalValues":
						$.isNumeric(r) && (r = Number(r)), $.isNumeric(u) && (u = Number(u)), n = t.DefaultDataValidator.createNumberValidator(f, r, u, !1);
						break;
					case "List":
						n = this._isFormula(r) ? t.DefaultDataValidator.createFormulaListValidator(r) : t.DefaultDataValidator.createListValidator(r);
						break;
					case "Date":
						n = t.DefaultDataValidator.createDateValidator(f, r, u);
						break;
					case "TextLength":
						n = t.DefaultDataValidator.createTextLengthValidator(f, r, u);
						break;
					case "Custom":
						n = t.DefaultDataValidator.createFormulaValidator(r);
						break
					}
					n.showInputMessage = this._showInputMessage.prop("checked"), n.inputTitle = this._inputTitle.val(), n.inputMessage = this._inputMessage.val(), n.showErrorMessage = this._showErrorMessage.prop("checked"), n.errorTitle = this._errorTitle.val(), n.errorMessage = this._errorMessage.val(), n.errorStyle = t.ErrorStyle[this._errorStyle.val()], n.ignoreBlank = this._ignoreBlank.prop("checked"), n.inCellDropdown = this._inCellDropDown.prop("checked"), n.comparisonOperator = t.ComparisonOperator[this._comparisonOperator.val()], i.actions.doAction("setDataValidation", i.wrapper.spread, n)
				}, u.prototype._isFormula = function(n) {
					return typeof n == "string" ? n.toString().charAt(0) == "=" : !1
				}, u.prototype._ValidValues = function() {
					var n = this._value1.val(),
						t = this._value2.val();
					if (this._value1.css("visibility") == "visible" && this._value2.css("visibility") == "hidden") if (!n) return i.MessageBox.show(i.res.dataValidationDialog.valueEmptyMessage, i.res.title, 1, 0), !1;
					if (this._value1.css("visibility") == "visible" && this._value2.css("visibility") == "visible") if (n && t) {
						if ($.isNumeric(n) && $.isNumeric(t)) if (Number(n) > Number(t)) return i.MessageBox.show(i.res.dataValidationDialog.minimumMaximumMessage, i.res.title, 1, 0), !1
					} else return i.MessageBox.show(i.res.dataValidationDialog.valueEmptyMessage, i.res.title, 1, 0), !1;
					return !0
				}, u
			}(i.BaseDialog), i.DataValidationDialog = ut, tt = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.groupDirectionDialog.settings,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting(), t.close(), i.wrapper.setFocusToSpread()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".group-direction-dialog", r)
				}
				return __extends(t, n), t.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet();
					this._element.find("#row-group-direction").prop("checked", n.rowRangeGroup.getDirection() == 1 ? !0 : !1), this._element.find("#column-group-direction").prop("checked", n.colRangeGroup.getDirection() == 1 ? !0 : !1)
				}, t.prototype._applySetting = function() {
					var n = i.wrapper.spread.getActiveSheet();
					n.isPaintSuspended(!0), n.rowRangeGroup.setDirection(this._element.find("#row-group-direction").prop("checked") ? 1 : 0), n.colRangeGroup.setDirection(this._element.find("#column-group-direction").prop("checked") ? 1 : 0), n.isPaintSuspended(!1)
				}, t
			}(i.BaseDialog), i.GroupDirectionDialog = tt, it = function(n) {
				function t() {
					var t = this,
						r = {
							resizable: !1,
							modal: !0,
							title: i.res.zoomDialog.title,
							width: 230,
							buttons: [{
								text: i.res.ok,
								click: function() {
									var u = t._element.find("input[name='zoom-ratio']:checked").val(),
										r, n;
									u == i.res.zoomDialog.fitSelection ? i.actions.doAction("zoomSelection", i.wrapper.spread) : (r = t._element.find("input[name='zoom-custom-ratio']").val(), n = parseFloat(r), n = n / 100, n ? i.actions.doAction("zoom", i.wrapper.spread, n) : i.MessageBox.show(i.res.zoomDialog.exception, i.res.title, 2)), t.close(), i.wrapper.setFocusToSpread(), i.ribbon.updateZoomToStatusBar()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".zoom-dialog", r)
				}
				return __extends(t, n), t.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet(),
						t = this._convertToPercent(n._zoomFactor.toString());
					this._element.find("input[type='text']").val(t), this._findFitRatio(n._zoomFactor)
				}, t.prototype._init = function() {
					var n = this,
						r = i.wrapper.spread.getActiveSheet(),
						f = r._zoomFactor,
						t = 4,
						u = this._convertToPercent(r._zoomFactor.toString());
					this._element.find("input[type='text']").val(u), this._findFitRatio(f), this._element.find("input[name='zoom-ratio']").bind("change", function() {
						var u = n._element.find("input[name='zoom-ratio']:checked").val(),
							r;
						u != i.res.zoomDialog.custom && (u == i.res.zoomDialog.fitSelection ? (r = i.actions._getPreferredZoomInfo(), r.zoom > t && (r.zoom = t), n._element.find("input[type='text']").val(n._convertToPercent(r.zoom.toString()))) : n._element.find("input[type='text']").val(u))
					}), this._element.find("input[name='zoom-custom-ratio']").bind("focus", function() {
						n._element.find("input[class='custom-ratio']").prop("checked", !0)
					})
				}, t.prototype._updateRatio = function() {
					var n = this._element.find("input[name='zoom-ratio']:checked").val();
					this._element.find("input[type='text']").val(n)
				}, t.prototype._convertToPercent = function(n) {
					return n.match(/^[0-9\.]*$/) ? parseFloat(n) * 100 : n.match(/^[0-9\.]*%$/) ? (n.substring(0, n.indexOf("%") - 1), n) : !1
				}, t.prototype._findFitRatio = function(n) {
					switch (n) {
					case 2:
						this._element.find("input[id='zoom-ratio1']").prop("checked", !0);
						break;
					case 1:
						this._element.find("input[id='zoom-ratio2']").prop("checked", !0);
						break;
					case.75:
						this._element.find("input[id='zoom-ratio3']").prop("checked", !0);
						break;
					case.5:
						this._element.find("input[id='zoom-ratio4']").prop("checked", !0);
						break;
					case.25:
						this._element.find("input[id='zoom-ratio5']").prop("checked", !0);
						break;
					default:
						this._element.find("input[class='custom-ratio']").prop("checked", !0)
					}
				}, t
			}(i.BaseDialog), i.ZoomDialog = it, function(n) {
				n[n.pie = 0] = "pie", n[n.area = 1] = "area", n[n.scatter = 2] = "scatter", n[n.compatible = 3] = "compatible", n[n.bullet = 4] = "bullet", n[n.spread = 5] = "spread", n[n.stacked = 6] = "stacked", n[n.hbar = 7] = "hbar", n[n.vbar = 8] = "vbar", n[n.variance = 9] = "variance", n[n.boxplot = 10] = "boxplot", n[n.cascade = 11] = "cascade", n[n.pareto = 12] = "pareto"
			}(i.FormulaSparklineType || (i.FormulaSparklineType = {})), bt = i.FormulaSparklineType, function(n) {
				n[n.cascade = 0] = "cascade", n[n.pareto = 1] = "pareto"
			}(i.MultiRangeFormulaType || (i.MultiRangeFormulaType = {})), pt = i.MultiRangeFormulaType, 
			rt = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.insertSparklineDialog.createSparklines,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".insert-sparkline-dialog", r)
				}
				return __extends(t, n), t.prototype._beforeOpen = function(n) {
					this._activeSheet = i.wrapper.spread.getActiveSheet(), this._formulaSparklineType = undefined, this._sparklineType = undefined;
					var t = this._element.find("#is-formula-sparkline"),
						u = t.parent(),
						r = n[0];
					n[1] !== undefined ? (this._formulaSparklineType = r, t.prop("checked", !0).attr("disabled", "disabled"), u.addClass("element-disabled")) : (this._sparklineType = r, t.prop("checked", !0).removeAttr("disabled"), u.removeClass("element-disabled")), this._element.find(".sparkline-data-range").val(i.CEUtility.parseRangeToExpString(this._activeSheet.getSelections()[0])), this._element.find(".sparkline-location-range").val("")
				}, t.prototype._applySetting = function() {
					var u = this,
						r = i.CEUtility.parseExpStringToRanges(this._element.find(".sparkline-data-range").val()),
						n = i.CEUtility.parseExpStringToRanges(this._element.find(".sparkline-location-range").val()),
						f;
					if (!r || r.length < 1 || !n || n.length < 1) {
						i.MessageBox.show(i.res.insertSparklineDialog.errorDataRangeMessage, i.res.title, 3, 0);
						return
					}
					f = this._element.find("#is-formula-sparkline").prop("checked");
					if (f) {
						var o = "",
							e = this._element.find(".sparkline-data-range").val(),
							t = "";
						if (this._sparklineType !== undefined) switch (this._sparklineType) {
						case 0:
							this._formulaSparklineType = 3, t = "line";
							break;
						case 1:
							this._formulaSparklineType = 3, t = "column";
							break;
						case 2:
							this._formulaSparklineType = 3, t = "winloss";
							break;
						default:
							break
						}
						i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
							type: this._formulaSparklineType,
							compatibleType: t,
							dataRange: e,
							locationRange: n[0]
						})
					} else i.actions.doAction("setDefaultSparkline", i.wrapper.spread, {
						dataRange: r[0],
						locationRange: n[0],
						sparklineType: this._sparklineType
					});
					i.ribbon.updateSparkline(), u.close(), u._updateFormulaBar(), i.wrapper.setFocusToSpread()
				}, t.prototype._updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this._activeSheet.getFormula(this._activeSheet.getActiveRowIndex(), this._activeSheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, t
			}(i.BaseDialog), i.InsertSparkLineDialog = rt, st = function(n) {
				function t() {
					var t = this,
						r = {
							width: "200",
							modal: !0,
							title: i.res.sparklineWeightDialog.sparklineWeight,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".sparkline-weight-dialog", r)
				}
				return __extends(t, n), t.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet(),
						t = n.getSparkline(n.getActiveRowIndex(), n.getActiveColumnIndex());
					t && this._element.find(".sparkline-weight").val(t.setting().lineWeight)
				}, t.prototype._applySetting = function() {
					var n = this._element.find(".sparkline-weight").val(),
						r = i.wrapper.spread.getActiveSheet(),
						t = this;
					$.isNumeric(n) ? (i.actions.doAction("setSparklineWeight", i.wrapper.spread, n), t.close(), i.wrapper.setFocusToSpread()) : i.MessageBox.show(i.res.sparklineWeightDialog.errorMessage, i.res.title, 2, 0)
				}, t
			}(i.BaseDialog), i.SparklineWeightDialog = st, vt = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.sparklineMarkerColorDialog.sparklineMarkerColor,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting(), t.close(), i.wrapper.setFocusToSpread()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".sparkline-marker-color-dialog", r)
				}
				return __extends(t, n), t.prototype._init = function() {
					var n = this;
					this._createColorPicker("sparkline-negative-point-color-frame", "sparkline-negative-point-color-span", "sparkline-negative-point-color-picker"), this._createColorPicker("sparkline-marker-point-color-frame", "sparkline-marker-point-color-span", "sparkline-marker-point-color-picker"), this._createColorPicker("sparkline-high-point-color-frame", "sparkline-high-point-color-span", "sparkline-high-point-color-picker"), this._createColorPicker("sparkline-low-point-color-frame", "sparkline-low-point-color-span", "sparkline-low-point-color-picker"), this._createColorPicker("sparkline-first-point-color-frame", "sparkline-first-point-color-span", "sparkline-first-point-color-picker"), this._createColorPicker("sparkline-last-point-color-frame", "sparkline-last-point-color-span", "sparkline-last-point-color-picker")
				}, t.prototype._createColorPicker = function(n, t, i) {
					var r = this;
					this._element.find("." + i).wijcolorpicker({
						valueChanged: function(n, i) {
							i.color === undefined ? r._element.find("." + t).css("background-color", "") : r._element.find("." + t).css("background-color", i.color)
						},
						choosedColor: function(t, i) {
							r._element.find("." + n).wijcomboframe("close")
						},
						openColorDialog: function(t, i) {
							r._element.find("." + n).wijcomboframe("close")
						}
					}), this._element.find("." + n).wijcomboframe()
				}, t.prototype._loadColor = function(n, t, r) {
					var f, u;
					$("." + r).wijcolorpicker({
						themeColors: i.wrapper.getThemeColors()
					}), f = i.wrapper.spread.getActiveSheet(), n ? (u = i.ColorHelper.parse(n, f.currentTheme().colors()), $("." + r).wijcolorpicker("option", "selectedItem", u), this._element.find("." + t).css("background-color", u.color)) : (this._element.find("." + t).css("background-color", "transparent"), $("." + r).wijcolorpicker("option", "selectedItem", null))
				}, t.prototype._saveColor = function(n, t) {
					n && (n.name ? i.actions.doAction(t, i.wrapper.spread, n.name) : i.actions.doAction(t, i.wrapper.spread, n.color))
				}, t.prototype._beforeOpen = function() {
					var t = i.wrapper.spread.getActiveSheet(),
						n = t.getSparkline(t.getActiveRowIndex(), t.getActiveColumnIndex()).setting();
					this._loadColor(n.negativeColor(), "sparkline-negative-point-color-span", "sparkline-negative-point-color-picker"), this._loadColor(n.markersColor(), "sparkline-marker-point-color-span", "sparkline-marker-point-color-picker"), this._loadColor(n.highMarkerColor(), "sparkline-high-point-color-span", "sparkline-high-point-color-picker"), this._loadColor(n.lowMarkerColor(), "sparkline-low-point-color-span", "sparkline-low-point-color-picker"), this._loadColor(n.firstMarkerColor(), "sparkline-first-point-color-span", "sparkline-first-point-color-picker"), this._loadColor(n.lastMarkerColor(), "sparkline-last-point-color-span", "sparkline-last-point-color-picker")
				}, t.prototype._applySetting = function() {
					this._saveColor($(".sparkline-negative-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineNegativeColor"), this._saveColor($(".sparkline-marker-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineMarkerColor"), this._saveColor($(".sparkline-high-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineHighColor"), this._saveColor($(".sparkline-low-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineLowColor"), this._saveColor($(".sparkline-first-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineFirstColor"), this._saveColor($(".sparkline-last-point-color-picker").wijcolorpicker("option", "selectedItem"), "setSparklineLastColor")
				}, t
			}(i.BaseDialog), i.SparklineMarkerColorDialog = vt, yt = function(n) {
				function t() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.createTableDialog.createTable,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting(), t.close(), i.wrapper.setFocusToSpread(), i.ribbon.updateTable()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".create-table-dialog", r)
				}
				return __extends(t, n), t.prototype._init = function() {
					var n = this
				}, t.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet();
					this._element.find(".table-source-range").val(i.CEUtility.parseRangeToExpString(n.getSelections()[0]))
				}, t.prototype._applySetting = function() {
					var n = i.CEUtility.parseExpStringToRanges(this._element.find(".table-source-range").val());
					n && n.length >= 1 && i.actions.doAction("createDefaultTable", i.wrapper.spread, n[0])
				}, t
			}(i.BaseDialog), i.CreateTableDialog = yt, ct = function(n) {
				function r() {
					var t = this,
						u = {
							width: 600,
							modal: !0,
							title: i.res.formatTableStyle.tableStyle,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._existTableStyleName() ? i.MessageBox.show(i.res.formatTableStyle.exception, i.res.title, 2) : (r._currentId++, t._storageStyle(), i.actions.isFileModified = !0, t.close(), i.wrapper.setFocusToSpread())
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".format-table-dialog", u)
				}
				return __extends(r, n), r.prototype._init = function() {
					this._tableStyle = null, this._firstColumnStripStyle = null, this._secondColumnStripStyle = null, this._firstRowStripStyle = null, this._secondRowStripStyle = null, this._createTableStyleName(), this._createTableElement(), this._createTablePreview(), this._createStripeSize(), this._attachEvent()
				}, r.prototype._createTableElement = function() {
					var r = this,
						t = i.res.tableElement,
						n;
					for (n in t) $("<option></option>").attr("strip-size", 1).val(n).text(t[n]).appendTo(this._element.find(".table-element-select"));
					this._element.find(".table-element-select").attr("size", 8), this._element.find(".table-element-select").get(0).selectedIndex = 0, this._tableStyleElement = this._element.find(".table-element-select").find("option:selected").val()
				}, r.prototype._createTableStyleName = function() {
					var n = i.res.formatTableStyle.tableStyle + " " + r._currentId;
					this._element.find(".table-style-input").val(n)
				}, r.prototype._createTablePreview = function() {
					var s = this,
						n, r, u, e, o, f;
					for (this._element.find(".pre-spread").wijspread(), this._tableFormatSpread = this._element.find(".pre-spread").wijspread("spread"), n = this._tableFormatSpread.getActiveSheet(), this._tableFormatSpread.showHorizontalScrollbar(!1), this._tableFormatSpread.showVerticalScrollbar(!1), this._tableFormatSpread.tabStripVisible(!1), this._tableFormatSpread.allowUserResize(!1), n.isPaintSuspended(!0), $(n._canvas).unbind("gcmousewheel.gcSheet"), n.currentTheme(i.wrapper.spread.getActiveSheet().currentTheme()), n.setGridlineOptions({
						showVerticalGridline: !1,
						showHorizontalGridline: !1
					}), n.colHeaderVisible = !1, n.rowHeaderVisible = !1, n.canUserDragDrop(!1), n.canUserDragFill(!1), n.selectionBackColor("transparent"), n.selectionBorderColor("transparent"), n.isProtected = !0, n.setColumnCount(8), n.setRowCount(9), r = 1; r < 8; r++) for (u = 1; u < 7; u++) n.setValue(r, u, "-"), n.getCell(r, u).vAlign($.sku.wijspread.VerticalAlign.center), n.getCell(r, u).hAlign($.sku.wijspread.HorizontalAlign.center);
					for (n.setRowHeight(0, 1), n.setRowHeight(8, 10), n.setColumnWidth(0, 3), n.setColumnWidth(7, 10), r = 1; r < 8; r++) n.setRowHeight(r, 14);
					for (u = 1; u < 7; u++) n.setColumnWidth(u, 17);
					if (n.getTables().length === 0) for (o = new t.TableStyle, e = n.addTable("table1", 1, 1, 6, 6, o), e.rowFilter().showFilterButton = !1, f = 0; f < 7; f++) e.setColumnName(f, "c" + f);
					n.isPaintSuspended(!1), setTimeout(function() {
						s._element.find(".pre-spread").wijspread("refresh")
					}, 0)
				}, r.prototype._createStripeSize = function() {
					for (var t = 8, n = 1; n <= t; n++) $("<option></option>").val(n).text(n).appendTo(this._element.find(".stripe-size-select"));
					this._element.find(".stripe-size-select").attr("size", 1), this._element.find(".stripe-size-select").get(0).selectedIndex = 0, this._element.find(".stripe-size-block").css("display", "none")
				}, r.prototype._attachEvent = function() {
					var n = this;
					this._element.find(".table-element-select").change(function() {
						var t = n._element.find(".stripe-size-block");
						n._tableStyleElement = n._element.find(".table-element-select").find("option:selected").val(), n._tableStyleElement.indexOf("Strip") === -1 ? t.css("display", "none") : t.css("display", "block"), n._syncStripSize()
					}), this._element.find(".stripe-size-select").change(function() {
						var t = n._element.find(".stripe-size-select").find("option:selected").val();
						n._element.find(".table-element-select").find("option:selected").attr("strip-size", t), n._updatePreview()
					}), this._element.find(".format-table-element").click(function() {
						var e = n._tableFormatSpread.getActiveSheet(),
							r = new t.Style;
						if (e.getTables().length !== 0) {
							var o = e.findTable(2, 2),
								s = n._element.find(".table-element-select").find("option:selected").val(),
								f = o.style(),
								u = f[s].call(f);
							u && (r.backColor = u.backColor, r.foreColor = u.foreColor, r.font = u.font, r.borderLeft = u.borderLeft, r.borderTop = u.borderTop, r.borderRight = u.borderRight, r.borderBottom = u.borderBottom)
						}
						this._formatDialog === undefined && (this._formatDialog = new i.FormatDialog), this._formatDialog.open("font", undefined, r, !0), this._formatDialog.selectTabOptions = {
							font: !0,
							border: !0,
							fill: !0
						}, this._formatDialog.setFormatDirectly(!1);
						$(this._formatDialog).on("okClicked", function(t, i) {
							var r = n._element.find(".table-element-select").find("option:selected").val();
							switch (r) {
							case "firstRowStripStyle":
								n._firstRowStripStyle = i;
								break;
							case "secondRowStripStyle":
								n._secondRowStripStyle = i;
								break;
							case "firstColumnStripStyle":
								n._firstColumnStripStyle = i;
								break;
							case "secondColumnStripStyle":
								n._secondColumnStripStyle = i;
								break;
							default:
								n._tableStyle = i;
								break
							}
							n._updatePreview()
						})
					}), this._element.find(".clear-table-element").click(function() {
						var t = n._tableFormatSpread.getActiveSheet();
						if (t.getTables().length !== 0) {
							t.isPaintSuspended(!0);
							var r = t.findTable(2, 2),
								u = n._element.find(".table-element-select").find("option:selected").val(),
								i = r.style();
							switch (u) {
							case "firstRowStripStyle":
								n._firstRowStripStyle = null;
								break;
							case "secondRowStripStyle":
								n._secondRowStripStyle = null;
								break;
							case "firstColumnStripStyle":
								n._firstColumnStripStyle = null;
								break;
							case "secondColumnStripStyle":
								n._secondColumnStripStyle = null;
								break;
							default:
								n._tableStyle = null;
								break
							}
							i[u].call(i, null), r.style(i), t.isPaintSuspended(!1)
						}
					})
				}, r.prototype._removeEvent = function() {
					this._element.find(".format-table-element").unbind("click"), this._element.find(".clear-table-element").unbind("click"), this._element.find(".table-element-select").unbind("change"), this._element.find(".stripe-size-select").unbind("change")
				}, r.prototype._syncStripSize = function() {
					var n, r = this._tableFormatSpread.getActiveSheet(),
						i = r.findTable(2, 2),
						f = parseInt(this._element.find(".stripe-size-select").find("option:selected").val()),
						u = this._element.find(".table-element-select").find("option:selected").val(),
						t;
					if (i !== null && this._element.find(".stripe-size-block").css("display") === "block") {
						t = i.style();
						switch (u) {
						case "firstRowStripStyle":
							n = t.firstRowStripSize();
							break;
						case "secondRowStripStyle":
							n = t.secondRowStripSize();
							break;
						case "firstColumnStripStyle":
							n = t.firstColumnStripSize();
							break;
						case "secondColumnStripStyle":
							n = t.secondColumnStripSize();
							break;
						default:
							break
						}
						n !== f && this._element.find(".stripe-size-select").val(n)
					}
				}, r.prototype._updatePreview = function() {
					var n, i, s = new t.TableStyle,
						o = this._element.find(".stripe-size-block"),
						u = this._tableFormatSpread.getActiveSheet(),
						e, r, f;
					u.isPaintSuspended(!0);
					if (u.getTables().length !== 0) n = u.findTable(2, 2);
					else for (n = u.addTable("table1", 1, 1, 6, 6, s), n.rowFilter().showFilterButton = !1, e = 0; e < 7; e++) n.setColumnName(e, "-");
					n.bandColumns(!0), n.bandRows(!0), n.showHeader(!0), n.showFooter(!0), n.highlightFirstColumn(!0), n.highlightLastColumn(!0), r = this._element.find(".table-element-select").find("option:selected").val();
					if (o.css("display") === "block") {
						f = parseInt(this._element.find(".stripe-size-select").find("option:selected").val());
						switch (r) {
						case "firstRowStripStyle":
							this._firstRowStripStyle && (i = this._getStyle(n, r, this._firstRowStripStyle), i.firstRowStripSize(f));
							break;
						case "secondRowStripStyle":
							this._secondRowStripStyle && (i = this._getStyle(n, r, this._secondRowStripStyle), i.secondRowStripSize(f));
							break;
						case "firstColumnStripStyle":
							this._firstColumnStripStyle && (i = this._getStyle(n, r, this._firstColumnStripStyle), i.firstColumnStripSize(f));
							break;
						case "secondColumnStripStyle":
							this._secondColumnStripStyle && (i = this._getStyle(n, r, this._secondColumnStripStyle), i.secondColumnStripSize(f));
							break;
						default:
							break
						}
					} else this._tableStyle && (i = this._getStyle(n, r, this._tableStyle));
					i && n.style(i), u.isPaintSuspended(!1)
				}, r.prototype._getStyle = function(n, i, r) {
					var e = new t.TableStyle,
						u = new t.TableStyleInfo,
						f;
					return r ? (f = n.style(), u = new t.TableStyleInfo, u.backColor = r.backColor, u.foreColor = r.foreColor, u.font = r.font, u.borderLeft = r.borderLeft, u.borderTop = r.borderTop, u.borderRight = r.borderRight, u.borderBottom = r.borderBottom, e = e[i].call(e, u), f[i].call(f, e[i]()), f) : null
				}, r.prototype._storageStyle = function() {
					var o = this._tableFormatSpread.getActiveSheet(),
						e = "custom",
						u = 1,
						i = o.findTable(1, 1),
						n, t, f;
					if (i) {
						n = this._element.find(".table-style-input").val(), t = i.style(), t.name(n);
						for (f in r.customTableStyle) u++;
						n = e + u + "-" + n, r.customTableStyle[n] = t
					}
				}, r.prototype._existTableStyleName = function() {
					var t = this._element.find(".table-style-input").val(),
						i = "custom",
						n;
					for (n in r.customTableStyle) {
						n = n.substring(n.indexOf("-") + 1);
						if (n === t) return !0
					}
					return !1
				}, r.prototype.refresh = function() {
					this._removeEvent(), this._element.find(".table-element-select").children().length !== 0 && this._element.find(".table-element-select").empty(), this._element.find(".pre-spread").children().length !== 0 && this._element.find(".pre-spread").empty(), this._element.find(".stripe-size-select").children().length !== 0 && this._element.find(".stripe-size-select").empty(), this._init()
				}, r._currentId = 1, r.customTableStyle = {}, r
			}(i.BaseDialog), i.FormatTableDialog = ct, ht = function(n) {
				function r() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.resizeTableDialog.title,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting(), t.close(), i.wrapper.setFocusToSpread()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".resize-table-dialog", r)
				}
				return __extends(r, n), r.prototype._beforeOpen = function() {
					var n = i.wrapper.spread.getActiveSheet(),
						r = n.findTable(n.getActiveRowIndex(), n.getActiveColumnIndex());
					r instanceof t.SheetTable && this._element.find(".table-source-range").val(i.CEUtility.parseRangeToExpString(r.range()))
				}, r.prototype._applySetting = function() {
					var n = i.CEUtility.parseExpStringToRanges(this._element.find(".table-source-range").val());
					n && n.length >= 1 && i.actions.doAction("resizeTable", i.wrapper.spread, {
						rowCount: n[0].rowCount,
						colCount: n[0].colCount
					})
				}, r
			}(i.BaseDialog), i.ResizeTableDialog = ht, at = function(n) {
				function r() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: "Style",
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.ok()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".cell-style-dialog", r)
				}
				return __extends(r, n), r.prototype._init = function() {
					this._element.find(".cell-style-format-setting").button(), this._attchEvent()
				}, r.prototype._beforeOpen = function() {
					this._setNameForStyle(), this._style = new t.Style
				}, r.prototype._setNameForStyle = function() {
					var n = "Style " + r._currentID;
					this._element.find(".cell-style-name").val(n)
				}, r.prototype.ok = function() {
					var n = this._element.find(".cell-style-name").val();
					this._isExisted(n) ? i.MessageBox.show(i.res.newCellStyleDialog.message, i.res.title, 2) : (this._style.name = n, r.existedCustomCellStyle[n.toLowerCase()] = this._style, r._currentID++, i.actions.isFileModified = !0, this.close(), i.wrapper.setFocusToSpread())
				}, r.prototype._isExisted = function(n) {
					if (n) {
						n = n.toLowerCase();
						if (r.existedCustomCellStyle.hasOwnProperty(n) || r.buildInCellStyle.hasOwnProperty(n)) return !0
					}
					return !1
				}, r.prototype._attchEvent = function() {
					var n = this;
					this._element.find(".cell-style-format-setting").click(function() {
						n._formatDialog === undefined && (n._formatDialog = new i.FormatDialog), n._formatDialog.open("numbers", undefined, n._style, !0), n._formatDialog.setFormatDirectly(!1);
						$(n._formatDialog).on("okClicked", function(t, i) {
							n._style = i
						})
					})
				}, r.buildInCellStyle = {}, r.existedCustomCellStyle = {}, r._currentID = 1, r
			}(i.BaseDialog), i.CellStyleDialog = at, k = function(n) {
				function t() {
					var t = this,
						r = {
							width: "450px",
							modal: !0,
							title: i.res.fileMenu.printerSettingDialogTitle,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.ok(), t.close()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".pdf-printer-setting-dialog", r)
				}
				return __extends(t, n), t.prototype._init = function() {
					this._printerRanges = [], this._element.find(".printer-range-setting-button").button(), this._initScalingSelect(), this._initDuplexSelect(), this._attachEvent()
				}, t.prototype._initScalingSelect = function() {
					var t = i.res.fileMenu.scalingType,
						r = this._element.find(".scaling-type"),
						n;
					for (n in t) $("<option></option>").val(n).text(t[n]).appendTo(r)
				}, t.prototype._initDuplexSelect = function() {
					var t = i.res.fileMenu.duplexMode,
						r = this._element.find(".duplex-mode"),
						n;
					for (n in t) $("<option></option>").val(n).text(t[n]).appendTo(r)
				}, t.prototype._changeActiveRow = function(n) {
					this._currentActiveRow && $(this._currentActiveRow).css("background-color", ""), this._currentActiveRow = n, $(this._currentActiveRow).css("background-color", "#98DDFB")
				}, t.prototype._attachEvent = function() {
					var n = this,
						t = this._element.find(".printer-range-table");
					$(".add-print-range").on("click", function() {
						var u = parseInt(n._element.find(".page-range-index").val(), 10),
							f = parseInt(n._element.find(".page-range-count").val(), 10),
							r;
						if (isNaN(u) || isNaN(f) || u < 0 || f <= 0) {
							i.MessageBox.show(i.res.fileMenu.addRangeException, i.res.title, 3, 0);
							return
						}
						r = $("<tr></tr>"), $("<td></td>").text(u).appendTo(r), $("<td></td>").text(f).appendTo(r), r.appendTo(t);
						$(r).on("click", function() {
							n._changeActiveRow(this)
						});
						n._printerRanges.push({
							index: u,
							count: f
						})
					});
					$(".remove-print-range").on("click", function() {
						var i, t, r;
						if (n._currentActiveRow) {
							i = $(n._currentActiveRow).find("td");
							if ($(i).length !== 2) return;
							var f = $(i[0]).text(),
								e = $(i[1]).text(),
								u = n._printerRanges,
								o = u.length;
							for (t = 0; t < o; t++) {
								r = u[t];
								if (r.index === parseInt(f, 10) && r.count === parseInt(e, 10)) {
									u.splice(t, 1);
									break
								}
							}
							$(n._currentActiveRow).remove()
						}
					})
				}, t.prototype.ok = function() {
					var r = $(".copies-number").find("option:selected").text(),
						u = $(".scaling-type").get(0).selectedIndex,
						n = $(".duplex-mode").get(0).selectedIndex,
						i = $(".papersource-by-pagesize").prop("checked");
					t.printerSettings = {
						scalingType: u,
						duplexMode: n,
						isPageSourceByPageSize: i,
						numberOfCopies: parseInt(r, 10),
						printRanges: this._printerRanges
					}
				}, t.printerSettings = {}, t
			}(i.BaseDialog), i.PDFPrinterDialog = k, lt = function(n) {
				function r() {
					var t = this,
						r = {
							width: "auto",
							modal: !0,
							title: i.res.formatComment.title,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.ok(), t.close()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".format-comment", r)
				}
				return __extends(r, n), r.prototype._init = function() {
					this._tabs = ["font", "alignment", "colors", "size", "protection", "properties"], this._defaultFontFamily = "Arial", this._defaultFontSize = "9pt", this._fillColorPicker = this._createColorPicker(this._element.find(".fill-color-picker"), {
						type: "comment-fill"
					}), this._lineColorPicker = this._createColorPicker(this._element.find(".line-color-picker"), {
						type: "comment-line"
					}), this._element.find(".color-transparent-slider").slider({
						range: "min",
						min: 0,
						max: 100
					}), this._attachEvent()
				}, r.prototype._beforeOpen = function(n) {
					var r, u, t;
					this._wijspread = i.wrapper.spread, r = this._wijspread.getActiveSheet();
					if (!r) return;
					n.length > 0 && (this._comment = r.getComment(n[0].row, n[0].col));
					if (!this._comment) return;
					this._element.find(".main-comment-tab").tabs(), this._element.find(".comment-tab-container li").removeClass("hidden"), u = 0, n.length > 0 && (t = n[0], t.showTabs && this._showTabs(t.showTabs), t.activeTab && (u = this._getTabIndexById(t.activeTab))), this._font = {}, this._alignment = {}, this._autoMaticSize = undefined, this._autoMaticPadding ? this._padding = {
						left: 0,
						right: 0,
						top: 0,
						bottom: 0
					} : (this._autoMaticPadding = !1, this._padding = this._getPadding()), this._lineStyle = {}, this._fillStyle = {}, this._size = {}, this._updateFont(), this._updateAlignment(), this._updateColorsAndLines(), this._updateSize(), this._updateProtection(), this._updateDynamic()
				}, r.prototype._getPadding = function() {
					var i = parseFloat(this._element.find("#padding-left").val()),
						r = parseFloat(this._element.find("#padding-right").val()),
						n = parseFloat(this._element.find("#padding-top").val()),
						t = parseFloat(this._element.find("#padding-bottom").val());
					return i = isNaN(i) ? 0 : i, r = isNaN(r) ? 0 : r, n = isNaN(n) ? 0 : n, t = isNaN(t) ? 0 : t, {
						left: i,
						right: r,
						top: n,
						bottom: t
					}
				}, r.prototype._attachEvent = function() {
					var n = this;
					this._element.find(".comment-font-picker").wijfontpicker({
						changed: function(t, i) {
							switch (i.name) {
							case "family":
								n._font.family = i.value;
								break;
							case "style":
								n._font.style = i.value;
								break;
							case "size":
								n._font.size = i.value;
								break;
							case "weight":
								n._font.weight = i.value;
								break;
							case "color":
								n._font.foreColor = i.value;
								break;
							case "underline":
								i.value ? n._font.textDecoration |= 1 : n._font.textDecoration &= 254;
								break;
							case "strikethrough":
								i.value ? n._font.textDecoration |= 2 : n._font.textDecoration &= 253;
								break
							}
						}
					}), this._element.find(".hAlign-select").change(function() {
						n._alignment.hAlign = $(this).val().toLowerCase()
					}), this._element.find("#auto-size").change(function() {
						n._autoMaticSize = $(this).prop("checked")
					}), this._element.find("#automatic-padding").change(function() {
						n._autoMaticPadding = $(this).prop("checked"), n._autoMaticPadding ? n._element.find(".padding-box").val("") : n._element.find(".padding-box").val("0"), n._padding = {
							left: 0,
							right: 0,
							top: 0,
							bottom: 0
						}
					}), this._element.find("#padding-left").spinner({
						min: 0,
						change: function() {
							n._padding.left = parseInt($(this).val()), n._autoMaticPadding = !1, n._updateCheckBoxValue(n._element.find("#automatic-padding"), n._autoMaticPadding)
						}
					}), this._element.find("#padding-right").spinner({
						min: 0,
						change: function() {
							n._padding.right = parseInt($(this).val()), n._autoMaticPadding = !1, n._updateCheckBoxValue(n._element.find("#automatic-padding"), n._autoMaticPadding)
						}
					}), this._element.find("#padding-top").spinner({
						min: 0,
						change: function() {
							n._padding.top = parseInt($(this).val()), n._autoMaticPadding = !1, n._updateCheckBoxValue(n._element.find("#automatic-padding"), n._autoMaticPadding)
						}
					}), this._element.find("#padding-bottom").spinner({
						min: 0,
						change: function() {
							n._padding.bottom = parseInt($(this).val()), n._autoMaticPadding = !1, n._updateCheckBoxValue(n._element.find("#automatic-padding"), n._autoMaticPadding)
						}
					}), this._element.find(".color-transparent-slider").slider({
						slide: function(t, i) {
							n._element.find("#colors-transparent-input").val(i.value);
							var r = parseFloat((i.value / 100).toFixed(2));
							n._fillStyle.transparency = r, n._element.find(".comment-fill").css("opacity", 1 - r)
						}
					}), this._element.find("#colors-transparent-input").spinner({
						min: 0,
						max: 100,
						change: function() {
							var t = parseInt($(this).val()),
								i;
							if (isNaN(t)) return;
							t < 0 ? (t = 0, $(this).val(t)) : t > 100 && (t = 100, $(this).val(t)), n._element.find(".color-transparent-slider").slider({
								value: t
							}), i = parseFloat((t / 100).toFixed(2)), n._fillStyle.transparency = i, n._element.find(".comment-fill").css("opacity", 1 - i)
						}
					}), this._element.find(".lines-style-select").change(function() {
						n._lineStyle.style = $(this).val()
					}), this._element.find(".lines-width").spinner({
						min: 0,
						change: function() {
							var t = parseFloat($(this).val());
							isNaN(t) || (n._lineStyle.width = t)
						}
					}), this._element.find("#comment-size-height").spinner({
						min: 0,
						change: function() {
							var t = parseFloat($(this).val());
							isNaN(t) || (n._size.height = t), n._autoMaticSize = !1
						}
					}), this._element.find("#comment-size-width").spinner({
						min: 0,
						change: function() {
							var t = parseFloat($(this).val());
							isNaN(t) || (n._size.width = t), n._autoMaticSize = !1
						}
					}), this._element.find("#comment-locked").change(function() {
						n._locked = $(this).prop("checked")
					}), this._element.find("#comment-lock-text").change(function() {
						n._lockText = $(this).prop("checked")
					}), this._element.find("input[type='radio'][name='dynamic-ratio']").change(function() {
						var t = n._element.find("input[type='radio'][name='dynamic-ratio']:checked");
						switch (t.attr("id")) {
						case "move-size":
							n._dynamicMove = !0, n._dynamicSize = !0;
							break;
						case "move-no-size":
							n._dynamicMove = !0, n._dynamicSize = !1;
							break;
						case "no-move-size":
							n._dynamicMove = !1, n._dynamicSize = !1;
							break
						}
					})
				}, r.prototype.ok = function() {
					this._setFont(), this._setAlignment(), this._setColorsAndLines(), this._setSize(), this._setProtection(), this._setDynamic()
				}, r.prototype._setFont = function() {
					var t, n, r;
					if (!this._comment) return;
					$.isEmptyObject(this._font) || (this._font.family && (n = this._comment.fontFamily(), n !== this._font.family && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "fontFamily",
						oldValue: n,
						newValue: this._font.family
					})), this._font.size && (n = parseFloat(this._comment.fontSize()), r = parseFloat(this._font.size), n !== r && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "fontSize",
						oldValue: n + "pt",
						newValue: this._font.size + "pt"
					})), this._font.weight && (n = this._comment.fontWeight(), this._font.weight !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "fontWeight",
						oldValue: n,
						newValue: this._font.weight
					})), this._font.style && (n = this._comment.fontStyle(), this._font.style !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "fontStyle",
						oldValue: n,
						newValue: this._font.style
					})), this._font.foreColor && (n = this._comment.foreColor(), t = this._font.foreColor, typeof t == "string" ? t !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "foreColor",
						oldValue: n,
						newValue: t
					}) : t.color !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "foreColor",
						oldValue: n,
						newValue: t.color
					})), this._font.textDecoration !== undefined && (n = this._comment.textDecoration(), r = this._font.textDecoration, r !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "textDecoration",
						oldValue: n,
						newValue: this._font.textDecoration
					})))
				}, r.prototype._updateFont = function() {
					var n, t;
					if (!this._comment) return;
					var u = this._comment.fontFamily(),
						e = this._comment.fontSize(),
						o = this._comment.fontWeight(),
						s = this._comment.fontStyle(),
						f = this._comment.foreColor(),
						r = this._comment.textDecoration();
					this._font.family = u ? u : this._defaultFontFamily, this._font.size = e ? e : this._defaultFontSize, this._font.weight = o ? o : undefined, this._font.style = s ? s : undefined, this._font.foreColor = f ? f : undefined, this._font.textDecoration = r ? r : 0, n = this._element.find(".comment-font-picker"), n.wijfontpicker("family", this._font.family), n.wijfontpicker("style", this._font.style ? this._font.style : "normal"), t = this._font.size, t.indexOf("px") !== -1 && (t = i.util.px2pt(parseInt(t))), n.wijfontpicker("size", parseInt(t)), n.wijfontpicker("weight", this._font.weight ? this._font.weight : "normal"), this._font.foreColor ? n.wijfontpicker("color", i.ColorHelper.parse(this._font.foreColor, i.wrapper.spread.getActiveSheet().currentTheme().colors())) : n.wijfontpicker("color", null), this._font.textDecoration & 1 ? n.wijfontpicker("underline", !0) : n.wijfontpicker("underline", !1), this._font.textDecoration & 2 ? n.wijfontpicker("strikethrough", !0) : n.wijfontpicker("strikethrough", !1)
				}, r.prototype._setAlignment = function() {
					var u, n, r;
					if (!this._comment) return;
					this._alignment.hAlign !== undefined && (n = this._comment.horizontalAlign(), u = t.HorizontalAlign[this._alignment.hAlign], u !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "horizontalAlign",
						oldValue: n,
						newValue: u
					})), this._autoMaticSize !== undefined && (n = this._comment.autoSize(), this._autoMaticSize !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "autoSize",
						oldValue: n,
						newValue: this._autoMaticSize
					})), this._autoMaticPadding !== undefined && (n = this._comment.padding(), r = this._autoMaticPadding ? null : $.isEmptyObject(this._padding) ? null : new t.Padding(this._padding.top, this._padding.right, this._padding.bottom, this._padding.left), this._isEqualPadding(n, r) || i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "padding",
						oldValue: n,
						newValue: r
					}))
				}, r.prototype._isEqualPadding = function(n, t) {
					if (!n && !t) return !0;
					if (n && t) if (n.left === t.left && n.right === t.right && n.top === t.top && n.bottom === t.bottom) return !0;
					return !1
				}, r.prototype._updateAlignment = function() {
					var i, r, n;
					if (!this._comment) return;
					i = this._comment.autoSize(), r = this._comment.horizontalAlign(), r !== undefined && this._updateComboBoxValue(this._element.find(".hAlign-select"), t.HorizontalAlign[r]), i !== undefined && this._updateCheckBoxValue(this._element.find("#auto-size"), i), n = this._comment.padding(), n ? (n.left !== undefined && (this._updateTextBoxValue(this._element.find("#padding-left"), n.left), this._autoMaticPadding = !1), n.right !== undefined && (this._updateTextBoxValue(this._element.find("#padding-right"), n.right), this._autoMaticPadding = !1), n.top !== undefined && (this._updateTextBoxValue(this._element.find("#padding-top"), n.top), this._autoMaticPadding = !1), n.bottom !== undefined && (this._updateTextBoxValue(this._element.find("#padding-bottom"), n.bottom), this._autoMaticPadding = !1)) : (this._autoMaticPadding = !0, this._element.find(".padding-box").val("")), this._updateCheckBoxValue(this._element.find("#automatic-padding"), this._autoMaticPadding)
				}, r.prototype._setColorsAndLines = function() {
					var o = this,
						t = this._comment,
						n, f, e, r, u;
					if (!t) return;
					n = o._fillStyle, n.color !== t.backColor() && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "backColor",
						oldValue: t.backColor(),
						newValue: n.color
					}), n.transparency !== 1 - t.opacity() && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "opacity",
						oldValue: t.opacity(),
						newValue: 1 - n.transparency
					}), $.isEmptyObject(this._lineStyle) || (f = this._comment.borderColor(), this._lineStyle.colorItem && (e = this._lineStyle.colorItem.color, e !== f && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "borderColor",
						oldValue: f,
						newValue: this._lineStyle.colorItem.color
					})), r = this._comment.borderStyle(), this._lineStyle.style && this._lineStyle.style !== r && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "borderStyle",
						oldValue: r,
						newValue: this._lineStyle.style
					}), u = this._comment.borderWidth(), this._lineStyle.width && this._lineStyle.width !== u && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "borderWidth",
						oldValue: u,
						newValue: this._lineStyle.width
					}))
				}, r.prototype._updateColorsAndLines = function() {
					var t, u, f, n, r, i;
					if (!this._comment) return;
					t = this._comment.backColor(), u = parseFloat((1 - this._comment.opacity()).toFixed(2)), this._fillStyle.color = t, this._fillStyle.transparency = u, this._fillColorPicker.wijcolorpicker("option", "selectedItem", t), this._element.find(".comment-fill").css("background-color", t), f = u * 100, this._element.find(".color-transparent-slider").slider("option", "value", f), this._element.find("#colors-transparent-input").val(f), n = this._comment.borderColor(), n && (this._lineColorPicker.wijcolorpicker("option", "selectedItem", n), this._element.find(".comment-line").css("background-color", n)), r = parseFloat(this._comment.borderWidth().toFixed(2)), isNaN(r) || this._element.find(".lines-width").val(r), i = this._comment.borderStyle(), i && this._element.find(".lines-style-select").val(i)
				}, r.prototype._setSize = function() {
					var t, n;
					if (!this._comment) return;
					this._size.height === undefined || this._autoMaticSize || (t = this._comment.height(), this._size.height !== t && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "height",
						oldValue: t,
						newValue: this._size.height
					})), this._size.width === undefined || this._autoMaticSize || (n = this._comment.width(), this._size.width !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "width",
						oldValue: n,
						newValue: this._size.width
					}))
				}, r.prototype._updateSize = function() {
					if (!this._comment) return;
					var t = this._comment.height(),
						n = this._comment.width();
					t !== undefined && this._element.find("#comment-size-height").val(t), n !== undefined && this._element.find("#comment-size-width").val(n)
				}, r.prototype._setProtection = function() {
					var n;
					if (!this._comment) return;
					this._locked !== undefined && (n = this._comment.locked(), this._locked !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "locked",
						oldValue: n,
						newValue: this._locked
					})), this._lockText !== undefined && (n = this._comment.lockText(), this._lockText !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "lockText",
						oldValue: n,
						newValue: this._lockText
					}))
				}, r.prototype._updateProtection = function() {
					if (!this._comment) return;
					var t = this._comment.locked(),
						n = this._comment.lockText();
					t !== undefined && this._element.find("#comment-locked").prop("checked", t), n !== undefined && this._element.find("#comment-lock-text").prop("checked", n)
				}, r.prototype._setDynamic = function() {
					var n;
					if (!this._comment) return;
					this._dynamicMove !== undefined && (n = this._comment.dynamicMove(), this._dynamicMove !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "dynamicMove",
						oldValue: n,
						newValue: this._dynamicMove
					})), this._dynamicSize !== undefined && (n = this._comment.dynamicSize(), this._dynamicSize !== n && i.actions.doAction("setCommentProperty", this._wijspread, {
						propertyName: "dynamicSize",
						oldValue: n,
						newValue: this._dynamicSize
					}))
				}, r.prototype._updateDynamic = function() {
					if (!this._comment) return;
					var n = this._comment.dynamicMove(),
						t = this._comment.dynamicSize();
					n && t ? this._element.find("#move-size").prop("checked", !0) : n && !t ? this._element.find("#move-no-size").prop("checked", !0) : n || this._element.find("#no-move-size").prop("checked", !0)
				}, r.prototype._showTabs = function(n) {
					var r = this._element.find(".comment-tab-container li"),
						t, i;
					r.addClass("hidden");
					for (t in n) n.hasOwnProperty(t) && (i = $.inArray(t, this._tabs), i !== -1 && n[t] && $(r[i]).removeClass("hidden"))
				}, r.prototype._createColorPicker = function(n, t) {
					if (!t.type) return;
					var r = this,
						i = $("<div></div>").addClass("colorpicker-container"),
						e = $("<span></span>").appendTo(i),
						o = $("<span></span>").addClass("colorpicker-preview " + t.type).appendTo(e),
						u = $("<div></div>").addClass("colorpicker-popup").appendTo(i),
						f = $("<div></div>").wijcolorpicker({
							valueChanged: function(n, i) {
								var u = i.color === undefined ? "transparent" : i.color;
								r._element.find("." + t.type).css("background-color", u);
								switch (t.type) {
								case "comment-fill":
									r._fillStyle.color = i.color;
									break;
								case "comment-line":
									r._lineStyle.colorItem = i;
									break
								}
							},
							choosedColor: function(n, t) {
								i.wijcomboframe("close")
							},
							openColorDialog: function(n, t) {
								i.wijcomboframe("close")
							}
						}).appendTo(u);
					return i.wijcomboframe().click(function() {
						i.wijcomboframe("open")
					}), n.append(i), f
				}, r.prototype._getTabIndexById = function(n) {
					switch (n) {
					case "font":
						return 0;
					case "alignment":
						return 1;
					case "colors":
						return 2;
					case "size":
						return 3;
					case "protection":
						return 4;
					case "properties":
						return 5
					}
					return 0
				}, r.prototype._updateCheckBoxValue = function(n, t) {
					var i = this._element.find(n);
					t === !0 ? (i.prop("indeterminate", !1), i.prop("checked", !0)) : t === !1 ? (i.prop("indeterminate", !1), i.prop("checked", !1)) : (i.prop("indeterminate", !0), i.prop("checked", !1))
				}, r.prototype._updateComboBoxValue = function(n, t) {
					var r = this._element.find(n);
					t !== i.BaseMetaObject.indeterminateValue ? r.val(t) : r.val(null)
				}, r.prototype._updateTextBoxValue = function(n, t) {
					var i = this._element.find(n);
					t !== undefined && t !== null ? i.val(t) : i.val(null)
				}, r.prototype._parseRGB = function(n) {
					var e = n.toLowerCase(),
						t, f, i, r, u;
					return e.indexOf("rgb") === -1 && e.indexOf("rgba") === -1 ? null : (t = n.split(","), t.length === 4 ? (i = t[0].substr(t[0].indexOf("(") + 1), r = t[1], u = t[2], f = t[3].substr(0, t[3].indexOf(")")), [parseFloat(i), parseFloat(r), parseFloat(u), parseFloat(f)]) : t.length === 3 ? (i = t[0].substr(t[0].indexOf("(") + 1), r = t[1], u = t[2].substr(0, t[2].indexOf(")")), [parseFloat(i), parseFloat(r), parseFloat(u)]) : null)
				}, r
			}(i.BaseDialog), i.FormatCommentDialog = lt, l = function(n) {
				function t() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.pieSparklineDialog.pieSparklineSetting,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".pie-sparkline-dialog", r)
				}
				return __extends(t, n), t.prototype._init = function() {
					this._commonWidgetClass = "pie-color-picker-widget", this._addPieColor(!0), this._defaultValue = {
						color: "#A9A9A9"
					}
				}, t.prototype._beforeOpen = function(n) {
					var v, y, e, u, l, h, t, f, a;
					this._sheet = i.wrapper.spread.getActiveSheet();
					if (!this._sheet) return;
					if (!n[0]) return;
					var s = n[0].row,
						c = n[0].col,
						o = i.util.parseFormulaSparkline(s, c);
					if (o.args) {
						v = o.args[0], y = i.util.unParseFormula(v, s, c), this._element.find(".pie-percentage").val(y), e = this._element.find(".pie-colors-list").children().length, u = o.args.length - 1;
						if (u > e) for (t = 0; t < u - e; t++) this._element.find(".add-pie-color").trigger("click");
						else if (u < e) for (t = 0; t < e - u; t++) l = this._element.find(".remove-pie-color"), l.length > 0 && $(l[0]).trigger("click");
						if (u === 0) h = i.ColorHelper.parse(this._defaultValue.color, this._sheet.currentTheme().colors()), $(".pie-color-picker1-widget").wijcolorpicker("option", "selectedItem", h), this._element.find(".pie-color-picker1").css("background-color", h.color);
						else for (t = 1; t <= u; t++) {
							f = null, a = r.parseColorExpression(o.args[t], s, c);
							try {
								f = i.ColorHelper.parse(a, this._sheet.currentTheme().colors())
							} catch (p) {}
							$(".pie-color-picker" + t + "-widget").wijcolorpicker("option", "selectedItem", f), this._element.find(".pie-color-picker" + t).css("background-color", f && f.color ? f.color : "transparent")
						}
					}
				}, t.prototype._updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this._sheet.getFormula(this._sheet.getActiveRowIndex(), this._sheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, t.prototype._applySetting = function() {
					var e = this._element.find(".pie-percentage").val(),
						o = this._element.find(".pie-colors-list").children().length,
						t = [],
						f, r, u, n, h, s;
					for (t.push(e), n = 0; n < o; n++) f = n + 1, r = $(".pie-color-picker" + f + "-widget").wijcolorpicker("option").selectedItem, (o !== 1 || r.baseColor !== this._defaultValue.color) && r && r.color && t.push(r.color);
					for (u = "=PIESPARKLINE(" + t[0], n = 1; n < t.length; n++) u += ',"' + t[n] + '"';
					u += ")", h = this._sheet.getActiveRowIndex(), s = this._sheet.getActiveColumnIndex(), i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						dataRange: e,
						type: 0
					}), this._updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, t.prototype._attachEvent = function() {
					var n = this;
					this._element.find(".add-pie-color").bind("click", function() {
						n._addPieColor()
					})
				}, t.prototype._addPieColor = function(n) {
					var e = this._element.find(".add-pie-color");
					n ? e.addClass("add-pie-color") : e.addClass("remove-pie-color").removeClass("add-pie-color");
					var s = this._element.find(".pie-colors-list"),
						t = this._getColorIndex(),
						r = $("<div>").addClass("pie-color").data("value", t).appendTo(s),
						c = $("<label>").text(i.res.pieSparklineDialog.color + " " + t).addClass("pie-color-column").appendTo(r),
						h = "pie-color-picker" + t,
						o = $("<div>").addClass("pie-color-picker").appendTo(r);
					this._createColorPicker(o, h);
					var f = $("<div>").addClass("add-pie-color").appendTo(r),
						l = $("<span>").addClass("ui-icon ui-icon-plus").appendTo(f),
						u = this;
					$(f).bind("click", function() {
						u._addPieColor()
					}), this._element.find(".remove-pie-color").unbind(), this._element.find(".remove-pie-color").bind("click", function(n) {
						u._removePieColor(n)
					})
				}, t.prototype._removePieColor = function(n) {
					var r = this,
						i = $(n.target).parents("div.pie-color"),
						u = i.data("value"),
						t = $(".pie-color-picker" + u + "-widget");
					t.length > 0 && (t.wijcolorpicker("destroy"), t.parents("div.color-picker-popup").remove()), i && i.length !== 0 && $(n.target).parents("div.pie-color").remove(), r._resetColorPicker()
				}, t.prototype._resetColorPicker = function() {
					for (var o = i.res.pieSparklineDialog.color, u = this._element.find(".pie-colors-list label"), s = u.length, f, t, n = 1; n <= s; n++) {
						$(u[n - 1]).text(o + " " + n);
						var e = $(u[n - 1]).parents("div.pie-color"),
							h = "pie-color-picker" + n,
							r = e.find("span.color-picker-preview");
						r.length > 0 && r.removeClass(r.attr("class")).addClass("color-picker-preview " + h), f = "pie-color-picker" + n + "-widget", t = $("." + this._commonWidgetClass), t.length > 0 && t[n - 1] && $(t[n - 1]).removeClass($(t[n - 1]).attr("class")).addClass(this._commonWidgetClass + " " + f), e.data("value", n)
					}
				}, t.prototype._createColorPicker = function(n, t) {
					var u = this,
						h = "color-picker-widget",
						s = t + "-widget",
						r = $("<div></div>").addClass("color-picker-container"),
						o = $("<span></span>").appendTo(r),
						c = $("<span></span>").addClass("color-picker-preview " + t).appendTo(o),
						e = $("<div></div>").addClass("color-picker-popup").appendTo(r),
						f = $("<div></div>").wijcolorpicker({
							valueChanged: function(n, t) {
								var i = t.color === undefined ? "transparent" : t.color;
								u._element.find(".pie-color-picker" + u._colorPickerIndex).css("background-color", i)
							},
							choosedColor: function(n, t) {
								r.wijcomboframe("close")
							},
							openColorDialog: function(n, t) {
								r.wijcomboframe("close")
							}
						}).addClass(this._commonWidgetClass + " " + s).appendTo(e);
					return f.wijcolorpicker("option", "themeColors", i.wrapper.getThemeColors()), r.wijcomboframe().click(function(n) {
						r.wijcomboframe("open");
						var t = $(n.target).parents("div.pie-color");
						u._colorPickerIndex = t.data("value")
					}), n.append(r), f
				}, t.prototype._getColorIndex = function() {
					for (var u = i.res.pieSparklineDialog.color, r = this._element.find(".pie-colors-list label"), t, n = r.length; n > 0; n--) {
						t = $(r[n - 1]).text();
						if (t === u + " " + n) return n + 1
					}
					return 1
				}, t
			}(i.BaseDialog), i.PieSparklineDialog = l, c = function(n) {
				function t() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.areaSparklineDialog.areaSparklineSetting,
							width: "500px",
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".area-sparkline-dialog", r)
				}
				return __extends(t, n), t.prototype._applySetting = function() {
					var f = this._element.find(".area-data-range").val(),
						l = this._element.find(".area-min-value").val(),
						c = this._element.find(".area-max-value").val(),
						h = this._element.find(".area-line1-value").val(),
						e = this._element.find(".area-line2-value").val(),
						t = $(".positive-color-picker").wijcolorpicker("option").selectedItem,
						u = "",
						n, r;
					t && t.color && t.baseColor !== this._defaultValue.colorPositive && (u = '"' + t.color + '"'), n = $(".negative-color-picker").wijcolorpicker("option").selectedItem, r = "", n && n.color && n.baseColor !== this._defaultValue.colorNegative && (r = '"' + n.color + '"');
					var o = [f, l, c, h, e, u, r],
						s = this._getFormula(o),
						v = this._sheet.getActiveRowIndex(),
						a = this._sheet.getActiveColumnIndex();
					i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: s,
						dataRange: f,
						type: 1
					}), this._updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, t.prototype._getFormula = function(n) {
					var i = n.length,
						r, t;
					while (i > 0 && n[i - 1] === "") i--;
					for (r = "", t = 0; t < i; t++) r += n[t], t !== i - 1 && (r += ",");
					return "=AREASPARKLINE(" + r + ")"
				}, t.prototype._beforeOpen = function(n) {
					var f, l, t, u, o, e;
					this._sheet = i.wrapper.spread.getActiveSheet();
					if (!this._sheet) return;
					if (!n[0]) return;
					var c = this._inputList,
						s = n[0].row,
						h = n[0].col,
						a = i.util.parseFormulaSparkline(s, h);
					if (a.args) {
						for (f = a.args, l = c.length, t = 0; t < l; t++) f[t] ? this._element.find(c[t]).val(i.util.unParseFormula(f[t], s, h)) : this._element.find(c[t]).val("");
						u = null, o = r.parseColorExpression(f[5], s, h), o || (o = this._defaultValue.colorPositive);
						try {
							u = i.ColorHelper.parse(o, this._sheet.currentTheme().colors())
						} catch (v) {}
						$(".positive-color-picker").wijcolorpicker("option", "selectedItem", u), this._element.find(".positive-color-span").css("background-color", u.color), e = r.parseColorExpression(f[6], s, h), e || (e = this._defaultValue.colorNegative);
						try {
							u = i.ColorHelper.parse(e, this._sheet.currentTheme().colors())
						} catch (v) {}
						$(".negative-color-picker").wijcolorpicker("option", "selectedItem", u), this._element.find(".negative-color-span").css("background-color", u.color)
					}
				}, t.prototype._init = function() {
					this._inputList = [".area-data-range", ".area-min-value", ".area-max-value", ".area-line1-value", ".area-line2-value"], this._createColorPicker(this._element.find(".positive-color"), "positive-color"), this._createColorPicker(this._element.find(".negative-color"), "negative-color"), this._defaultValue = {
						colorPositive: "#787878",
						colorNegative: "#CB0000"
					}
				}, t.prototype._updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this._sheet.getFormula(this._sheet.getActiveRowIndex(), this._sheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, t.prototype._createColorPicker = function(n, t) {
					var s = this,
						r = $("<div></div>").addClass("area-color-picker-container"),
						h = $("<span></span>").appendTo(r),
						u = t + "-span",
						c = $("<span></span>").addClass("area-color-picker-preview " + u).appendTo(h),
						f = $("<div></div>").addClass("area-color-picker-popup").appendTo(r),
						o = t + "-picker",
						e = $("<div></div>").wijcolorpicker({
							valueChanged: function(n, t) {
								var i = t.color === undefined ? "transparent" : t.color;
								s._element.find("." + u).css("background-color", i)
							},
							choosedColor: function(n, t) {
								r.wijcomboframe("close")
							},
							openColorDialog: function(n, t) {
								r.wijcomboframe("close")
							}
						}).addClass(o).appendTo(f);
					e.wijcolorpicker("option", "themeColors", i.wrapper.getThemeColors()), r.wijcomboframe().click(function(n) {
						r.wijcomboframe("open")
					}), n.append(r)
				}, t
			}(i.BaseDialog), i.AreaSparklineDialog = c, v = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.scatterSparklineDialog.scatterSparklineSetting,
							width: "600px",
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".scatter-sparkline-dialog", r), t._colorValueChanged = "ColorValueChanged"
				}
				return __extends(u, n), u.prototype._init = function() {
					this._inputList = [".scatter-points1", ".scatter-points2", ".scatter-minx", ".scatter-maxx", ".scatter-miny", ".scatter-maxy", ".scatter-hline", ".scatter-vline", ".scatter-xmin-zone", ".scatter-xmax-zone", ".scatter-ymin-zone", ".scatter-ymax-zone"], this._createColorPicker(this._element.find(".scatter-color1"), "scatter-color1"), this._createColorPicker(this._element.find(".scatter-color2"), "scatter-color2"), this._defaultValue = {
						tags: !1,
						drawSymbol: !0,
						drawLines: !1,
						dash: !1,
						color1: "#969696",
						color2: "#CB0000"
					}, this._attachEvent()
				}, u.prototype._applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t != undefined && t != null ? t + "," : ",";
					n = this._removeContinuousComma(n), u = "=SCATTERSPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						dataRange: this._paraPool[0],
						type: 2
					}), this._updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u.prototype._removeContinuousComma = function(n) {
					var t = n.length;
					while (t > 0 && n[t - 1] === ",") t--;
					return n.substr(0, t)
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".scatter-points1").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".scatter-points2").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".scatter-minx").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".scatter-maxx").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".scatter-miny").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".scatter-maxy").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), t.find(".scatter-hline").keyup(function() {
						n._paraPool[6] = $(this).val()
					}), t.find(".scatter-vline").keyup(function() {
						n._paraPool[7] = $(this).val()
					}), t.find(".scatter-xmin-zone").keyup(function() {
						n._paraPool[8] = $(this).val()
					}), t.find(".scatter-xmax-zone").keyup(function() {
						n._paraPool[9] = $(this).val()
					}), t.find(".scatter-ymin-zone").keyup(function() {
						n._paraPool[10] = $(this).val()
					}), t.find(".scatter-ymax-zone").keyup(function() {
						n._paraPool[11] = $(this).val()
					}), t.find(".scatter-tags").change(function() {
						n._paraPool[12] = $(this).prop("checked")
					}), t.find(".scatter-drawsymbol").change(function() {
						n._paraPool[13] = $(this).prop("checked")
					}), t.find(".scatter-drawlines").change(function() {
						n._paraPool[14] = $(this).prop("checked")
					}), $(".scatter-point1").bind(this._colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[15] = '"' + r + '"'
					}), $(".scatter-point2").bind(this._colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "tranparent";
						n._paraPool[16] = '"' + r + '"'
					}), t.find(".scatter-dash").change(function() {
						n._paraPool[17] = $(this).prop("checked")
					})
				}, u.prototype._beforeOpen = function(n) {
					var f, s;
					this._paraPool = [], this._sheet = i.wrapper.spread.getActiveSheet();
					if (!this._sheet) return;
					if (!n[0]) return;
					var c = n[0].row,
						h = n[0].col,
						w = i.util.parseFormulaSparkline(c, h);
					if (w.args) {
						var u = w.args,
							p = this._inputList,
							d = p.length;
						for (f = 0; f < d; f++) s = "", u[f] && (s = i.util.unParseFormula(u[f], c, h)), this._element.find(p[f]).val(s), this._paraPool.push(s);
						var y = u[12] instanceof t.Calc.Expressions.BooleanExpression ? u[12].value : null,
							l = u[13] instanceof t.Calc.Expressions.BooleanExpression ? u[13].value : null,
							v = u[14] instanceof t.Calc.Expressions.BooleanExpression ? u[14].value : null,
							e = r.parseColorExpression(u[15], c, h),
							o = r.parseColorExpression(u[16], c, h),
							a = u[17] instanceof t.Calc.Expressions.BooleanExpression ? u[17].value : null,
							k = e ? '"' + e + '"' : null,
							b = o ? '"' + o + '"' : null;
						this._updateCheckBox("scatter-tags", y !== null ? y : this._defaultValue.tags), this._updateCheckBox("scatter-drawsymbol", l !== null ? l : this._defaultValue.drawSymbol), this._updateCheckBox("scatter-drawlines", v !== null ? v : this._defaultValue.drawLines), this._updateCheckBox("scatter-dash", a !== null ? a : this._defaultValue.dash), this._updateColorPicker("scatter-color1", e !== null ? e : this._defaultValue.color1), this._updateColorPicker("scatter-color2", o !== null ? o : this._defaultValue.color2), this._paraPool.push(y, l, v, k, b, a)
					}
				}, u.prototype._updateColorPicker = function(n, t) {
					var u = $("." + n + "-picker"),
						r;
					if (u.length === 0) return;
					if (t) {
						r = null;
						try {
							r = i.ColorHelper.parse(t, this._sheet.currentTheme().colors())
						} catch (f) {}
						r && r.color && (u.wijcolorpicker("option", "selectedItem", r), this._element.find("." + n + "-span").css("background-color", r.color))
					} else u.wijcolorpicker("option", "selectedItem", null), this._element.find("." + n + "-span").css("background-color", "transparent")
				}, u.prototype._createColorPicker = function(n, t) {
					var u = this,
						r = $("<div></div>").addClass("scatter-color-picker-container"),
						h = $("<span></span>").appendTo(r),
						f = t + "-span",
						c = $("<span></span>").addClass("scatter-color-picker-preview " + f).appendTo(h),
						s = $("<div></div>").addClass("scatter-color-picker-popup").appendTo(r),
						e = t + "-picker",
						o = $("<div></div>").wijcolorpicker({
							valueChanged: function(n, t) {
								var i = t.color === undefined ? "transparent" : t.color;
								n.target.className === "scatter-color1-picker" ? u._paraPool[15] = '"' + i + '"' : u._paraPool[16] = '"' + i + '"', u._element.find("." + f).css("background-color", i), $(this).trigger(u._colorValueChanged, {
									value: t
								})
							},
							choosedColor: function(n, t) {
								r.wijcomboframe("close")
							},
							openColorDialog: function(n, t) {
								r.wijcomboframe("close")
							}
						}).addClass(e).appendTo(s);
					o.wijcolorpicker("option", "themeColors", i.wrapper.getThemeColors()), r.wijcomboframe().click(function(n) {
						r.wijcomboframe("open")
					}), n.append(r)
				}, u.prototype._updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this._sheet.getFormula(this._sheet.getActiveRowIndex(), this._sheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, u.prototype._updateTextBox = function(n, t) {
					var i = this._element.find("." + n);
					t !== undefined ? i.val(t) : i.val("")
				}, u.prototype._updateCheckBox = function(n, t, i) {
					t !== undefined ? this._element.find("." + n).prop("checked", t) : this._element.find("." + n).prop("checked", i)
				}, u
			}(i.BaseDialog), i.ScatterSparklineDialog = v, a = function(u) {
				function f() {
					var n = this,
						t = {
							modal: !0,
							title: i.res.compatibleSparklineDialog.compatibleSparklineSetting,
							width: "580px",
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									n._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									n.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					u.call(this, "./dialogs/dialogs2.html", ".compatible-sparkline-dialog", t)
				}
				return __extends(f, u), f.prototype._beforeOpen = function(n) {
					var t, o;
					this._sparklineName = "", this._sparklineSetting = {}, this._sheet = i.wrapper.spread.getActiveSheet();
					if (!this._sheet) return;
					if (!n[0]) return;
					var f = n[0].row,
						e = n[0].col,
						u = i.util.parseFormulaSparkline(f, e);
					u && (u.fn.name && (this._sparklineName = u.fn.name), u.args && u.args.length > 0 && (t = u.args, this._updateTextBox("com-data", i.util.unParseFormula(t[0], f, e)), this._updateSelect("com-data-orientation", t[1].value), t[2] ? this._updateTextBox("com-date-axis-data", i.util.unParseFormula(t[2], f, e)) : this._updateTextBox("com-date-axis-data"), t[3] ? this._updateSelect("com-date-axis-orientation", t[3].value) : this._updateSelect("com-date-axis-orientation"), o = r.parseColorExpression(t[4], f, e), o && (this._sparklineSetting = this._parseSetting(o)), this._updateSparklineSetting(this._sparklineSetting)))
				}, f.prototype._parseSetting = function(n) {
					var r = {},
						h = !1,
						s = !0,
						f = "",
						e = "",
						o, c, t, u, i;
					if (n) {
						for (n = n.substr(1, n.length - 2), o = 0, c = n.length; o < c; o++) t = n.charAt(o), t === ":" ? s = !1 : t !== "," || h ? t !== "'" && t !== '"' && (t === "(" ? h = !0 : t === ")" && (h = !1), s ? f += t : e += t) : (r[f] = e, f = "", e = "", s = !0);
						f && (r[f] = e);
						for (u in r) i = r[u], i !== null && typeof i != "undefined" && (i.toUpperCase() === "TRUE" ? r[u] = !0 : i.toUpperCase() === "FALSE" ? r[u] = !1 : !isNaN(i) && isFinite(i) && (r[u] = parseFloat(i)))
					}
					return r
				}, f.prototype._reset = function() {
					this._updateSelect("com-display-emptycell-as", -1), this._updateCheckBox("com-display-hidden", this._defaultSetting.displayHidden), this._updateCheckBox("com-show-first", this._defaultSetting.showFirst), this._updateCheckBox("com-show-last", this._defaultSetting.showLast), this._updateCheckBox("com-show-high", this._defaultSetting.showHigh), this._updateCheckBox("com-show-low", this._defaultSetting.showLow), this._updateCheckBox("com-show-negative", this._defaultSetting.showNegative), this._updateCheckBox("com-show-markers", this._defaultSetting.showMarkers), this._updateCheckBox("com-show-markers", this._defaultSetting.showMarkers), this._updateCheckBox("com-right-to-left", this._defaultSetting.rightToLeft), this._updateCheckBox("com-display-xaxis", this._defaultSetting.displayXAxis), this._updateSelect("com-min-axis-type", -1), this._updateSelect("com-max-axis-type", -1), this._updateTextBox("com-manual-max", ""), this._updateTextBox("com-manual-min", "")
				}, f.prototype._updateSparklineSetting = function(n) {
					var t;
					if (!n) return;
					this._updateSelect("com-display-emptycell-as", n.displayEmptyCellsAs !== undefined ? n.displayEmptyCellsAs : n.deca), this._updateCheckBox("com-display-hidden", n.displayHidden !== undefined ? n.displayHidden : n.dh, this._defaultSetting.displayHidden), this._updateCheckBox("com-show-first", n.showFirst !== undefined ? n.showFirst : n.sf, this._defaultSetting.showFirst), this._updateCheckBox("com-show-last", n.showLast !== undefined ? n.showLast : n.slast, this._defaultSetting.showLast), this._updateCheckBox("com-show-high", n.showHigh !== undefined ? n.showHigh : n.sh, this._defaultSetting.showHigh), this._updateCheckBox("com-show-low", n.showLow !== undefined ? n.showLow : n.slow, this._defaultSetting.showLow), this._updateCheckBox("com-show-negative", n.showNegative !== undefined ? n.showNegative : n.sn, this._defaultSetting.showNegative), this._updateCheckBox("com-show-markers", n.showMarkers !== undefined ? n.showMarkers : n.sm, this._defaultSetting.showMarkers), this._updateCheckBox("com-right-to-left", n.rightToLeft !== undefined ? n.rightToLeft : n.rtl, this._defaultSetting.rightToLeft), this._updateCheckBox("com-display-xaxis", n.displayXAxis !== undefined ? n.displayXAxis : n.dxa, this._defaultSetting.displayXAxis), this._updateSelect("com-min-axis-type", n.minAxisType !== undefined ? n.minAxisType : n.minat), this._updateSelect("com-max-axis-type", n.maxAxisType !== undefined ? n.maxAxisType : n.maxat), this._updateTextBox("com-manual-max", n.manualMax !== undefined ? n.manualMax : n.mmax), this._updateTextBox("com-manual-min", n.manualMin !== undefined ? n.manualMin : n.mmin), t = this._element.find(".com-min-axis-type").val(), this._updateManual(t, "com-manual-min-label", "com-manual-min"), t = this._element.find(".com-max-axis-type").val(), this._updateManual(t, "com-manual-max-label", "com-manual-max")
				}, f.prototype._updateSelect = function(n, t) {
					var i = this._element.find("." + n).get(0);
					t !== undefined ? (t < i.length || (t = i.length - 1), i.selectedIndex = t) : i.selectedIndex = -1
				}, f.prototype._updateTextBox = function(n, t) {
					var i = this._element.find("." + n);
					t !== undefined ? i.val(t) : i.val("")
				}, f.prototype._updateCheckBox = function(n, t, i) {
					t !== undefined ? (t = t ? !0 : !1, this._element.find("." + n).prop("checked", t)) : this._element.find("." + n).prop("checked", i)
				}, f.prototype._updateManual = function(n, t, i) {
					var u = this._element.find("." + i),
						r = this._element.find("." + t);
					n !== "custom" ? (u.addClass("manual-disable").attr("disabled", "disabled"), r.addClass("manual-disable")) : (u.removeClass("manual-disable").removeAttr("disabled"), r.removeClass("manual-disable"))
				}, f.prototype._init = function() {
					this._element.find(".com-color-setting").button(), this._defaultSetting = {
						displayEmptyCellsAs: 0,
						rightToLeft: !1,
						displayHidden: !1,
						displayXAxis: !1,
						showFirst: !1,
						showHigh: !1,
						showLast: !1,
						showLow: !1,
						showNegative: !1,
						showMarkers: !1,
						manualMax: 0,
						manualMin: 0,
						maxAxisType: 0,
						minAxisType: 0
					}, this._shortDic = {
						axiscolor: "ac",
						firstmarkercolor: "fmc",
						highmarkercolor: "hmc",
						lastmarkercolor: "lastmc",
						lowmarkercolor: "lowmc",
						markerscolor: "mc",
						negativecolor: "nc",
						seriescolor: "sc",
						displayemptycellsas: "deca",
						righttoleft: "rtl",
						displayhidden: "dh",
						displayxaxis: "dxa",
						showfirst: "sf",
						showhigh: "sh",
						showlast: "slast",
						showlow: "slow",
						shownegative: "sn",
						showmarkers: "sm",
						manualmax: "mmax",
						manualmin: "mmin",
						maxaxistype: "maxat",
						minaxistype: "minat",
						lineweight: "lw"
					}, this._styleSettingDialog || (this._styleSettingDialog = new e), this._attachEvent()
				}, f.prototype._attachEvent = function() {
					var n = this,
						i = n._element;
					i.find(".com-color-setting").bind("click", function() {
						n._styleSettingDialog.open(n._sparklineSetting)
					}), $(this._styleSettingDialog).bind("compatibleStyleOkClick", function(t, i) {
						n._sparklineSetting = $.extend({}, n._sparklineSetting, i)
					}), i.find(".com-display-emptycell-as").change(function() {
						var i = n._sparklineSetting,
							r = t.EmptyValueStyle[n._element.find(".com-display-emptycell-as").val()];
						i.deca !== undefined ? i.deca = r : i.displayEmptyCellsAs = r
					}), i.find(".com-display-hidden").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-display-hidden").prop("checked");
						t.dh !== undefined ? t.dh = i : t.displayHidden = i
					}), i.find(".com-show-first").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-first").prop("checked");
						t.sf !== undefined ? t.sf = i : t.showFirst = i
					}), i.find(".com-show-last").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-last").prop("checked");
						t.slast !== undefined ? t.slast = i : t.showLast = i
					}), i.find(".com-show-high").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-high").prop("checked");
						t.sh !== undefined ? t.sh = i : t.showHigh = i
					}), i.find(".com-show-low").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-low").prop("checked");
						t.slow ? t.slow = i : t.showLow = i
					}), i.find(".com-show-negative").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-negative").prop("checked");
						t.sn !== undefined ? t.sn = i : t.showNegative = i
					}), i.find(".com-show-markers").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-show-markers").prop("checked");
						t.sm !== undefined ? t.sm = i : t.showMarkers = i
					}), i.find(".com-min-axis-type").change(function() {
						var r = n._sparklineSetting,
							i = n._element.find(".com-min-axis-type").val();
						r.minat !== undefined ? r.minat = t.SparklineAxisMinMax[i] : r.minAxisType = t.SparklineAxisMinMax[i], n._updateManual(i, "com-manual-min-label", "com-manual-min")
					}), i.find(".com-max-axis-type").change(function() {
						var r = n._sparklineSetting,
							i = n._element.find(".com-max-axis-type").val();
						r.maxat !== undefined ? r.maxat = t.SparklineAxisMinMax[i] : r.maxAxisType = t.SparklineAxisMinMax[i], n._updateManual(i, "com-manual-max-label", "com-manual-max")
					}), i.find(".com-manual-max").keyup(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-manual-max").val();
						t.mmax !== undefined ? t.mmax = i : t.manualMax = i
					}), i.find(".com-manual-min").keyup(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-manual-min").val();
						t.mmin !== undefined ? t.mmin = i : t.manualMin = i
					}), i.find(".com-right-to-left").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-right-to-left").prop("checked");
						t.rtl !== undefined ? t.rtl = i : t.rightToLeft = i
					}), i.find(".com-display-xaxis").change(function() {
						var t = n._sparklineSetting,
							i = n._element.find(".com-display-xaxis").prop("checked");
						t.dxa !== undefined ? t.dxa = i : t.displayXAxis = i
					})
				}, f.prototype._applySetting = function() {
					var e = this._element.find(".com-data").val(),
						s = t.DataOrientation[this._element.find(".com-data-orientation").val()],
						h = this._element.find(".com-date-axis-data").val(),
						o = t.DataOrientation[this._element.find(".com-date-axis-orientation").val()],
						u, f, r, n;
					o === undefined && (o = ""), u = [];
					for (n in this._sparklineSetting) this._sparklineSetting[n] !== undefined && this._sparklineSetting[n] !== "" && u.push(n + ":" + this._sparklineSetting[n]);
					f = "", u.length > 0 && (f = '"{' + u.join(",") + '}"'), r = "", r = f !== "" ? "=" + this._sparklineName + "(" + e + "," + s + "," + h + "," + o + "," + f + ")" : o !== "" ? "=" + this._sparklineName + "(" + e + "," + s + "," + h + "," + o + ")" : h !== "" ? "=" + this._sparklineName + "(" + e + "," + s + "," + h + ")" : "=" + this._sparklineName + "(" + e + "," + s + ")";
					if (r.length >= 250) {
						u = [];
						for (n in this._sparklineSetting) this._sparklineSetting[n] !== undefined && this._sparklineSetting[n] !== "" && u.push(this._shortenFormula(n, !0) + ":" + this._shortenFormula(this._sparklineSetting[n])), f = '"{' + u.join(",") + '}"', r = "=" + this._sparklineName + "(" + e + "," + s + "," + h + "," + o + "," + f + ")"
					}
					i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: r,
						type: 3,
						dataRange: e
					}), this._updateFormulaBar(), this._reset(), this.close(), i.wrapper.setFocusToSpread()
				}, f.prototype._shortenFormula = function(t, i) {
					var o, u, e, f, r;
					if (i) return o = this._shortDic[t.toLowerCase()], o ? o : t;
					else if (t === !0) return 1;
					else if (t === !1) return 0;
					else if (t.length > 9) {
						t = t.toLowerCase();
						if (t.indexOf("rgb") !== -1) {
							u = n.spread.util.parseColorString(t);
							if (u && u.length > 0) for (e = "#", f = 0; f < u.length; f++) r = u[f].toString(16), r.length === 1 && (r = "0" + r), e += r;
							if (e !== "#") return e
						}
					}
					return t
				}, f.prototype._updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this._sheet.getFormula(this._sheet.getActiveRowIndex(), this._sheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, f
			}(i.BaseDialog), i.CompatibleSparklineDialog = a, e = function(n) {
				function t() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.compatibleSparklineDialog.styleSetting,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t._applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".compatible-style-setting", r)
				}
				return __extends(t, n), t.prototype._beforeOpen = function(n) {
					this._colorSets = {}, this._sheet = i.wrapper.spread.getActiveSheet();
					if (n[0] !== undefined) {
						var t = n[0];
						this._setting = n[0], this._updateColorPicker("axis", (t.axisColor ? t.axisColor : t.ac) || this._defaultValue.axis), this._updateColorPicker("series", (t.seriesColor ? t.seriesColor : t.sc) || this._defaultValue.series), this._updateColorPicker("markers", (t.markersColor ? t.markersColor : t.mc) || this._defaultValue.markers), this._updateColorPicker("negative", (t.negativeColor ? t.negativeColor : t.nc) || this._defaultValue.negativePoints), this._updateColorPicker("first-marker", (t.firstMarkerColor ? t.firstMarkerColor : t.fmc) || this._defaultValue.firstPoint), this._updateColorPicker("high-marker", (t.highMarkerColor ? t.highMarkerColor : t.hmc) || this._defaultValue.highPoint), this._updateColorPicker("last-marker", (t.lastMarkerColor ? t.lastMarkerColor : t.lastmc) || this._defaultValue.lastPoint), this._updateColorPicker("low-marker", (t.lowMarkerColor ? t.lowMarkerColor : t.lowmc) || this._defaultValue.lowPoint), this._updateTextBox("compatible-line-weight", t.lineWeight || t.lw)
					}
				}, t.prototype._init = function() {
					this._defaultValue = {
						negativePoints: "#A52A2A",
						markers: "#244062",
						highPoint: "#0000FF",
						lowPoint: "#0000FF",
						firstPoint: "#95B3D7",
						lastPoint: "#95B3D7",
						series: "#244062",
						axis: "#000000"
					}, this._shortDic = {
						axiscolor: "ac",
						firstmarkercolor: "fmc",
						highmarkercolor: "hmc",
						lastmarkercolor: "lastmc",
						lowmarkercolor: "lowmc",
						markerscolor: "mc",
						negativecolor: "nc",
						seriescolor: "sc",
						lineweight: "lw"
					}, this._createColorPicker(this._element.find(".compatible-axis"), "axis", "axisColor"), this._createColorPicker(this._element.find(".compatible-series"), "series", "seriesColor"), this._createColorPicker(this._element.find(".compatible-markers"), "markers", "markersColor"), this._createColorPicker(this._element.find(".compatible-negative"), "negative", "negativeColor"), this._createColorPicker(this._element.find(".compatible-first-marker"), "first-marker", "firstMarkerColor"), this._createColorPicker(this._element.find(".compatible-high-marker"), "high-marker", "highMarkerColor"), this._createColorPicker(this._element.find(".compatible-last-marker"), "last-marker", "lastMarkerColor"), this._createColorPicker(this._element.find(".compatible-low-marker"), "low-marker", "lowMarkerColor"), this._attachEvent()
				}, t.prototype._updateColorPicker = function(n, t) {
					var u = $("." + n + "-picker"),
						r;
					if (u.length === 0) return;
					if (t) {
						r = null;
						try {
							r = i.ColorHelper.parse(t, this._sheet.currentTheme().colors())
						} catch (f) {}
						r && r.color && (u.wijcolorpicker("option", "selectedItem", r), this._element.find("." + n + "-span").css("background-color", r.color))
					} else u.wijcolorpicker("option", "selectedItem", null), this._element.find("." + n + "-span").css("background-color", "transparent")
				}, t.prototype._updateTextBox = function(n, t) {
					t !== undefined ? this._element.find("." + n).val(t) : this._element.find("." + n).val("")
				}, t.prototype._attachEvent = function() {
					var n = this;
					this._element.find(".compatible-line-weight").keyup(function() {
						var t = n._element.find(".compatible-line-weight").val();
						n._setting.lineWeight !== undefined ? n._colorSets.lineWeight = t : n._colorSets.lw = t
					})
				}, t.prototype._reset = function() {
					this._updateColorPicker("axis", this._defaultValue.axis), this._updateColorPicker("series", this._defaultValue.series), this._updateColorPicker("markers", this._defaultValue.markers), this._updateColorPicker("negative", this._defaultValue.negativePoints), this._updateColorPicker("first-marker", this._defaultValue.firstPoint), this._updateColorPicker("high-marker", this._defaultValue.highPoint), this._updateColorPicker("last-marker", this._defaultValue.lastPoint), this._updateColorPicker("low-marker", this._defaultValue.lowPoint), this._updateTextBox("compatible-line-weight")
				}, t.prototype._applySetting = function() {
					$(this).trigger("compatibleStyleOkClick", this._colorSets), this._reset(), this.close()
				}, t.prototype._createColorPicker = function(n, t, r) {
					var f = this,
						u = $("<div></div>").addClass("compatible-color-picker-container"),
						c = $("<span></span>").appendTo(u),
						e = t + "-span",
						l = $("<span></span>").addClass("compatible-color-picker-preview " + e).appendTo(c),
						h = $("<div></div>").addClass("compatible-color-picker-popup").appendTo(u),
						o = t + "-picker",
						s = $("<div></div>").wijcolorpicker({
							valueChanged: function(n, t) {
								var u = t.color === undefined ? "transparent" : t.color,
									i;
								f._element.find("." + e).css("background-color", u), i = f._shortDic[r.toLowerCase()], f._setting[i] !== undefined ? f._colorSets[i] = u : f._colorSets[r] = u
							},
							choosedColor: function(n, t) {
								u.wijcomboframe("close")
							},
							openColorDialog: function(n, t) {
								u.wijcomboframe("close")
							}
						}).addClass(o).appendTo(h);
					s.wijcolorpicker("option", "themeColors", i.wrapper.getThemeColors()), u.wijcomboframe().click(function(n) {
						u.wijcomboframe("open")
					}), n.append(u)
				}, t
			}(i.BaseDialog), i.CompatibleStyleSettingDialog = e, u = function(n) {
				function t(t, r, u) {
					n.call(this, t, r, u), this.sheet = i.wrapper.spread.getActiveSheet(), this.colorValueChanged = "ColorValueChanged"
				}
				return __extends(t, n), t.prototype.updateTextBox = function(n, t) {
					var i = this._element.find("." + n);
					t !== undefined ? i.val(t) : i.val("")
				}, t.prototype.updateColorPicker = function(n, t, r) {
					var f = $("." + n),
						u;
					if (f.length === 0) return;
					if (r) {
						u = null;
						try {
							u = i.ColorHelper.parse(r, this.sheet.currentTheme().colors())
						} catch (e) {}
						u && u.color && (f.wijcolorpicker("option", "selectedItem", u), this._element.find("." + t).css("background-color", u.color))
					} else f.wijcolorpicker("option", "selectedItem", null), this._element.find("." + t).css("background-color", "transparent")
				}, t.prototype.updateCheckBox = function(n, t, i) {
					t !== undefined && t !== null ? this._element.find("." + n).prop("checked", t) : this._element.find("." + n).prop("checked", i)
				}, t.prototype.updateSelect = function(n, t) {
					var i = this._element.find("." + n);
					t !== undefined && i.val(t)
				}, t.prototype.removeContinuousComma = function(n) {
					var t = n.length;
					while (t > 0 && n[t - 1] === ",") t--;
					return n.substr(0, t)
				}, t.prototype.updateFormulaBar = function() {
					var t = $("#formulaBarText"),
						n;
					t.length > 0 && (n = this.sheet.getFormula(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex()), n && (n = "=" + n, t.text(n)))
				}, t.prototype.createColorPicker = function(n, t, r) {
					var u = this;
					this._element.find("." + r).wijcolorpicker({
						valueChanged: function(n, i) {
							i.color === undefined ? u._element.find("." + t).css("background-color", "transparent") : u._element.find("." + t).css("background-color", i.color), $(this).trigger(u.colorValueChanged, {
								value: i
							})
						},
						choosedColor: function(t, i) {
							u._element.find("." + n).wijcomboframe("close")
						},
						openColorDialog: function(t, i) {
							u._element.find("." + n).wijcomboframe("close")
						}
					}), this._element.find("." + r).wijcolorpicker("option", "themeColors", i.wrapper.getThemeColors()), this._element.find("." + n).wijcomboframe()
				}, t
			}(i.BaseDialog), i.SparklineExBaseDialog = u, h = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.bulletSparklineDialog.bulletSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".bullet-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._beforeOpen = function() {
					this._paraPool = [], this._updateBulletDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype._updateBulletDialog = function(n, u) {
					var o = i.util.parseFormulaSparkline(n, u),
						f, c;
					if (o && o.args) {
						f = o.args;
						if (f && f.length > 0) {
							var v = i.util.unParseFormula(f[0], n, u),
								a = i.util.unParseFormula(f[1], n, u),
								y = i.util.unParseFormula(f[2], n, u),
								w = i.util.unParseFormula(f[3], n, u),
								p = i.util.unParseFormula(f[4], n, u),
								s = i.util.unParseFormula(f[5], n, u),
								h = i.util.unParseFormula(f[6], n, u),
								e = r.parseColorExpression(f[7], n, u),
								l = f[8] instanceof t.Calc.Expressions.BooleanExpression ? f[8].value : null;
							this.updateTextBox("bullet-measure", v), this.updateTextBox("bullet-target", a), this.updateTextBox("bullet-maxi", y), this.updateTextBox("bullet-good", w), this.updateTextBox("bullet-bad", p), this.updateTextBox("bullet-forecast", s), this.updateTextBox("bullet-tickunit", h), this.updateColorPicker("bullet-color-picker", "bullet-color-span", e ? e : this._defaultValue.colorScheme), this.updateCheckBox("bullet-vertical", l, this._defaultValue.vertical), c = e ? '"' + e + '"' : null, this._paraPool = [v, a, y, w, p, s, h, c, l]
						}
					} else this._reset()
				}, u.prototype._reset = function() {
					this.updateTextBox("bullet-measure", ""), this.updateTextBox("bullet-target", ""), this.updateTextBox("bullet-maxi", ""), this.updateTextBox("bullet-good", ""), this.updateTextBox("bullet-bad", ""), this.updateTextBox("bullet-forecast", ""), this.updateTextBox("bullet-tickunit", ""), this.updateColorPicker("bullet-color-picker", "bullet-color-span", this._defaultValue.colorScheme), this.updateCheckBox("bullet-vertical", this._defaultValue.vertical)
				}, u.prototype._init = function() {
					this._defaultValue = {
						vertical: !1,
						colorScheme: "#A0A0A0"
					}, this.createColorPicker("bullet-color-frame", "bullet-color-span", "bullet-color-picker"), this._attchEvent()
				}, u.prototype._attchEvent = function() {
					var n = this,
						t = this._element;
					t.find(".bullet-measure").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".bullet-target").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".bullet-maxi").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".bullet-forecast").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), t.find(".bullet-good").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".bullet-bad").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".bullet-tickunit").keyup(function() {
						n._paraPool[6] = $(this).val()
					}), $(".bullet-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[7] = '"' + r + '"'
					}), t.find(".bullet-vertical").change(function() {
						n._paraPool[8] = $(this).prop("checked")
					})
				}, u.prototype.applySetting = function() {
					for (var e = this._paraPool, n = "", r, o, f, u = 0; u < e.length; u++) r = e[u], n += r !== undefined && r !== null ? r + "," : ",";
					n = this.removeContinuousComma(n), o = "=BULLETSPARKLINE(" + n + ")", f = this.sheet.getSelections(), i.actions.doAction("setFormulaSparklineWithoutDatarange", i.wrapper.spread, {
						formula: o,
						type: 4,
						locationRange: f.length > 0 ? f[0] : new t.Range(0, 0, 1, 1)
					}), i.ribbon.updateSparkline(), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u
			}(u), i.BulletSparklineDialog = h, s = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.spreadSparklineDialog.spreadSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".spread-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultValue = {
						showAverage: !1,
						style: 4,
						colorScheme: "#646464",
						vertical: !1
					}, this.createColorPicker("spread-color-frame", "spread-color-span", "spread-color-picker"), this._attachEvent()
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this._updateSpreadDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".spread-points").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".spread-style").change(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".spread-scale-start").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".spread-scale-end").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), $(".spread-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[5] = '"' + r + '"'
					}), t.find(".spread-show-average").change(function() {
						n._paraPool[1] = $(this).prop("checked")
					}), t.find(".spread-vertical").change(function() {
						n._paraPool[6] = $(this).prop("checked")
					})
				}, u.prototype._updateSpreadDialog = function(n, u) {
					var s = i.util.parseFormulaSparkline(n, u),
						f, c;
					if (s && s.args) {
						f = s.args;
						if (f && f.length > 0) {
							var a = i.util.unParseFormula(f[0], n, u),
								l = f[1] instanceof t.Calc.Expressions.BooleanExpression ? f[1].value : null,
								y = i.util.unParseFormula(f[2], n, u),
								v = i.util.unParseFormula(f[3], n, u),
								o = f[4] ? i.util.unParseFormula(f[4], n, u) : null,
								e = r.parseColorExpression(f[5], n, u),
								h = f[6] instanceof t.Calc.Expressions.BooleanExpression ? f[6].value : null;
							this.updateTextBox("spread-points", a), this.updateTextBox("spread-scale-start", y), this.updateTextBox("spread-scale-end", v), this.updateColorPicker("spread-color-picker", "spread-color-span", e ? e : this._defaultValue.colorScheme), this.updateCheckBox("spread-show-average", l, this._defaultValue.showAverage), this.updateCheckBox("spread-vertical", h, this._defaultValue.vertical), this.updateSelect("spread-style", o ? o : this._defaultValue.style), c = e ? '"' + e + '"' : null, this._paraPool = [a, l, y, v, o, c, h]
						} else this._reset()
					}
				}, u.prototype.applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t !== undefined && t !== null ? t + "," : ",";
					n = this.removeContinuousComma(n), u = "=SPREADSPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						type: 5,
						dataRange: this._paraPool[0]
					}), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u.prototype._reset = function() {
					this.updateTextBox("spread-points", ""), this.updateTextBox("spread-scale-start", ""), this.updateTextBox("spread-scale-end", ""), this.updateColorPicker("spread-color-picker", "spread-color-span", this._defaultValue.colorScheme), this.updateCheckBox("spread-show-average", this._defaultValue.showAverage), this.updateCheckBox("spread-vertical", this._defaultValue.vertical), this.updateSelect("spread-style", this._defaultValue.style)
				}, u
			}(u), i.SpreadSparklineDialog = s, d = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.stackedSparklineDialog.stackedSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".stacked-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultValue = {
						color: "#646464",
						vertical: !1,
						textOrientation: 0
					}, this.createColorPicker("stacked-color-frame", "stacked-color-span", "stacked-color-picker"), this._attachEvent()
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = this._element;
					t.find(".stacked-points").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".stacked-color-range").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".stacked-label-range").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".stacked-maximum").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".stacked-target-red").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".stacked-target-green").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), t.find(".stacked-target-blue").keyup(function() {
						n._paraPool[6] = $(this).val()
					}), t.find(".stacked-target-yellow").keyup(function() {
						n._paraPool[7] = $(this).val()
					}), t.find(".stacked-highlight-position").keyup(function() {
						n._paraPool[9] = $(this).val()
					}), t.find(".stacked-vertical").change(function() {
						n._paraPool[10] = $(this).prop("checked")
					}), t.find(".stacked-text-orientation").change(function() {
						n._paraPool[11] = $(this).val()
					}), t.find(".stacked-text-size").keyup(function() {
						n._paraPool[12] = $(this).val()
					}), $(".stacked-color-picker").bind(n.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[8] = '"' + r + '"'
					})
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateStackedDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype.applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t !== undefined && t !== null ? t + "," : ",";
					n = this.removeContinuousComma(n), u = "=STACKEDSPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						type: 6,
						dataRange: this._paraPool[0]
					}), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u.prototype._updateStackedDialog = function(n, u) {
					var s = i.util.parseFormulaSparkline(n, u),
						f, l;
					if (s && s.args) {
						f = s.args;
						if (f && f.length > 0) {
							var w = i.util.unParseFormula(f[0], n, u),
								p = i.util.unParseFormula(f[1], n, u),
								y = i.util.unParseFormula(f[2], n, u),
								b = i.util.unParseFormula(f[3], n, u),
								g = i.util.unParseFormula(f[4], n, u),
								d = i.util.unParseFormula(f[5], n, u),
								k = i.util.unParseFormula(f[6], n, u),
								c = i.util.unParseFormula(f[7], n, u),
								e = r.parseColorExpression(f[8], n, u),
								h = i.util.unParseFormula(f[9], n, u),
								v = f[10] instanceof t.Calc.Expressions.BooleanExpression ? f[10].value : null,
								o = i.util.unParseFormula(f[11], n, u),
								a = i.util.unParseFormula(f[12], n, u);
							this.updateTextBox("stacked-points", w), this.updateTextBox("stacked-color-range", p), this.updateTextBox("stacked-label-range", y), this.updateTextBox("stacked-maximum", b), this.updateTextBox("stacked-target-red", g), this.updateTextBox("stacked-target-green", d), this.updateTextBox("stacked-target-blue", k), this.updateTextBox("stacked-target-yellow", c), this.updateColorPicker("stacked-color-picker", "stacked-color-span", e ? e : this._defaultValue.color), this.updateTextBox("stacked-highlight-position", h), this.updateCheckBox("stacked-vertical", v, this._defaultValue.vertical), this.updateSelect("stacked-text-orientation", o ? o : this._defaultValue.textOrientation), this.updateTextBox("stacked-text-size", a), l = e ? '"' + e + '"' : null, this._paraPool = [w, p, y, b, g, d, k, c, l, h, v, o, a]
						}
					} else this._reset()
				}, u.prototype._reset = function() {
					this.updateTextBox("stacked-points", ""), this.updateTextBox("stacked-color-range", ""), this.updateTextBox("stacked-label-range", ""), this.updateTextBox("stacked-maximum", ""), this.updateTextBox("stacked-target-red", ""), this.updateTextBox("stacked-target-green", ""), this.updateTextBox("stacked-target-blue", ""), this.updateTextBox("stacked-target-yellow", ""), this.updateColorPicker("stacked-color-picker", "stacked-color-span", this._defaultValue.color), this.updateTextBox("stacked-highlight-position", ""), this.updateCheckBox("stacked-vertical", this._defaultValue.vertical), this.updateTextBox("stacked-text-orientation", this._defaultValue.textOrientation), this.updateTextBox("stacked-text-size", "")
				}, u
			}(u), i.StackedSparklineDialog = d, f = function(n) {
				function u(t, i, r) {
					n.call(this, t, i, r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultColor = "gray", this.createColorPicker("barbase-color-frame", "barbase-color-span", "barbase-color-picker"), this._attachEvent()
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateBarBaseSparklineDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype._updateBarBaseSparklineDialog = function(n, t) {
					var e = i.util.parseFormulaSparkline(n, t),
						f, o, u, s;
					e && e.args ? (f = e.args, f && f.length > 0 && (o = i.util.unParseFormula(f[0], n, t), u = r.parseColorExpression(f[1], n, t), this.updateTextBox("barbase-value", o), this.updateColorPicker("barbase-color-picker", "barbase-color-span", u === null ? this._defaultColor : u), s = u ? '"' + u + '"' : null, this._paraPool = [o, s])) : this._reset()
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = this._element;
					t.find(".barbase-value").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), $(".barbase-color-picker").bind(n.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[1] = '"' + r + '"'
					})
				}, u.prototype._reset = function() {
					this.updateTextBox("barbase-value", ""), this.updateColorPicker("barbase-color-picker", "barbase-color-span", this._defaultColor)
				}, u.prototype.applySetting = function(n) {
					for (var s = this._paraPool, r = "", u, o, e, f = 0; f < s.length; f++) u = s[f], r += u !== undefined && u !== null ? u + "," : ",";
					r = this.removeContinuousComma(r), n === 7 ? o = "=HBARSPARKLINE(" + r + ")" : n === 8 && (o = "=VBARSPARKLINE(" + r + ")"), e = this.sheet.getSelections(), i.actions.doAction("setFormulaSparklineWithoutDatarange", i.wrapper.spread, {
						formula: o,
						type: n,
						locationRange: e.length > 0 ? e[0] : new t.Range(0, 0, 1, 1)
					}), i.ribbon.updateSparkline(), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u
			}(u), i.BarBaseSparklineDialog = f, nt = function(n) {
				function t() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.barbaseSparklineDialog.hbarSparklineSetting,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting(7)
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".barbase-sparkline-dialog", r)
				}
				return __extends(t, n), t
			}(f), i.HbarSparklineDialog = nt, g = function(n) {
				function t() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.barbaseSparklineDialog.vbarSparklineSetting,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting(8)
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".barbase-sparkline-dialog", r)
				}
				return __extends(t, n), t
			}(f), i.VbarSparklineDialog = g, p = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.variSparklineDialog.variSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".vari-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultValue = {
						legend: !1,
						colorPositive: "green",
						colorNegative: "red",
						vertical: !1
					}, this.createColorPicker("vari-negative-color-frame", "vari-negative-color-span", "vari-negative-color-picker"), this.createColorPicker("vari-positive-color-frame", "vari-positive-color-span", "vari-positive-color-picker"), this._attachEvent()
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateVariSparklineDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype._updateVariSparklineDialog = function(n, u) {
					var s = i.util.parseFormulaSparkline(n, u),
						f, l, c;
					if (s && s.args) {
						f = s.args;
						if (f && f.length > 0) {
							var p = i.util.unParseFormula(f[0], n, u),
								y = i.util.unParseFormula(f[1], n, u),
								v = i.util.unParseFormula(f[2], n, u),
								k = i.util.unParseFormula(f[3], n, u),
								b = i.util.unParseFormula(f[4], n, u),
								w = i.util.unParseFormula(f[5], n, u),
								h = f[6] instanceof t.Calc.Expressions.BooleanExpression ? f[6].value : null,
								e = r.parseColorExpression(f[7], n, u),
								o = r.parseColorExpression(f[8], n, u),
								a = f[9] instanceof t.Calc.Expressions.BooleanExpression ? f[9].value : null;
							this.updateTextBox("vari-variance", p), this.updateTextBox("vari-reference", y), this.updateTextBox("vari-mini", v), this.updateTextBox("vari-maxi", k), this.updateTextBox("vari-mark", b), this.updateTextBox("vari-tickunit", w), this.updateColorPicker("vari-positive-color-picker", "vari-positive-color-span", e ? e : this._defaultValue.colorPositive), this.updateColorPicker("vari-negative-color-picker", "vari-negative-color-span", o ? o : this._defaultValue.colorNegative), this.updateCheckBox("vari-legend", h), this.updateCheckBox("vari-vertical", a), l = e ? '"' + e + '"' : null, c = o ? '"' + o + '"' : null, this._paraPool = [p, y, v, k, b, w, h, l, c, a]
						}
					} else this._reset()
				}, u.prototype._reset = function() {
					this.updateTextBox("vari-variance", ""), this.updateTextBox("vari-reference", ""), this.updateTextBox("vari-mini", ""), this.updateTextBox("vari-maxi", ""), this.updateTextBox("vari-mark", ""), this.updateTextBox("vari-tickunit", ""), this.updateColorPicker("vari-positive-color-picker", "vari-positive-color-span", this._defaultValue.colorPositive), this.updateColorPicker("vari-negative-color-picker", "vari-negative-color-span", this._defaultValue.colorNegative), this.updateCheckBox("vari-legend", this._defaultValue.legend), this.updateCheckBox("vari-vertical", this._defaultValue.vertical)
				}, u.prototype.applySetting = function() {
					for (var e = this._paraPool, n = "", r, o, f, u = 0; u < e.length; u++) r = e[u], n += r !== undefined && r !== null ? r + "," : ",";
					n = this.removeContinuousComma(n), o = "=VARISPARKLINE(" + n + ")", f = this.sheet.getSelections(), i.actions.doAction("setFormulaSparklineWithoutDatarange", i.wrapper.spread, {
						formula: o,
						type: 4,
						locationRange: f.length > 0 ? f[0] : new t.Range(0, 0, 1, 1)
					}), i.ribbon.updateSparkline(), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".vari-variance").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".vari-reference").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".vari-mini").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".vari-maxi").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".vari-mark").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".vari-tickunit").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), $(".vari-positive-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[7] = '"' + r + '"'
					}), $(".vari-negative-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[8] = '"' + r + '"'
					}), t.find(".vari-legend").change(function() {
						n._paraPool[6] = $(this).prop("checked")
					}), t.find(".vari-vertical").change(function() {
						n._paraPool[9] = $(this).prop("checked")
					})
				}, u
			}(u), i.VariSparklineDialog = p, y = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.boxplotSparklineDialog.boxplotSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".boxplot-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultValue = {
						boxplotClass: "5ns",
						style: 0,
						colorScheme: "#D2D2D2",
						vertical: !1,
						showAverage: !1
					}, this.createColorPicker("boxplot-color-frame", "boxplot-color-span", "boxplot-color-picker"), this._attachEvent()
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateBoxPlotSparklineDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex())
				}, u.prototype._updateBoxPlotSparklineDialog = function(n, u) {
					var l = i.util.parseFormulaSparkline(n, u),
						f, y, v;
					if (l && l.args) {
						f = l.args;
						if (f && f.length > 0) {
							var e = this._defaultValue,
								w = i.util.unParseFormula(f[0], n, u),
								o = f[1] instanceof t.Calc.Expressions.StringExpression ? f[1].value : null,
								a = f[2] instanceof t.Calc.Expressions.BooleanExpression ? f[2].value : null,
								p = i.util.unParseFormula(f[3], n, u),
								b = i.util.unParseFormula(f[4], n, u),
								d = i.util.unParseFormula(f[5], n, u),
								k = i.util.unParseFormula(f[6], n, u),
								s = r.parseColorExpression(f[7], n, u),
								h = f[8] ? i.util.unParseFormula(f[8], n, u) : null,
								c = f[9] instanceof t.Calc.Expressions.BooleanExpression ? f[9].value : null;
							this.updateTextBox("boxplot-points", w), this.updateSelect("boxplot-class", o === null ? e.boxplotClass : o), this.updateTextBox("boxplot-scale-start", p), this.updateTextBox("boxplot-scale-end", b), this.updateTextBox("boxplot-acceptable-start", d), this.updateTextBox("boxplot-acceptable-end", k), this.updateColorPicker("boxplot-color-picker", "boxplot-color-span", s === null ? e.colorScheme : s), this.updateSelect("boxplot-style", h === null ? e.style : h), this.updateCheckBox("boxplot-vertical", c === null ? e.vertical : c), this.updateCheckBox("boxplot-show-average", a === null ? e.showAverage : a), y = o ? '"' + o + '"' : null, v = s ? '"' + s + '"' : null, this._paraPool = [w, y, a, p, b, d, k, v, h, c]
						} else this._reset()
					}
				}, u.prototype._reset = function() {
					this.updateTextBox("boxplot-points", ""), this.updateSelect("boxplot-class", this._defaultValue.boxplotClass), this.updateTextBox("boxplot-scale-start", ""), this.updateTextBox("boxplot-scale-end", ""), this.updateTextBox("boxplot-acceptable-start", ""), this.updateTextBox("boxplot-acceptable-end", ""), this.updateColorPicker("boxplot-color-picker", "boxplot-color-span", this._defaultValue.colorScheme), this.updateSelect("boxplot-style", this._defaultValue.style), this.updateCheckBox("boxplot-vertical", this._defaultValue.vertical), this.updateCheckBox("boxplot-show-average", this._defaultValue.showAverage)
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".boxplot-points").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".boxplot-class").change(function() {
						n._paraPool[1] = '"' + $(this).val() + '"'
					}), t.find(".boxplot-scale-start").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".boxplot-scale-end").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".boxplot-acceptable-start").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), t.find(".boxplot-acceptable-end").keyup(function() {
						n._paraPool[6] = $(this).val()
					}), $(".boxplot-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[7] = '"' + r + '"'
					}), t.find(".boxplot-style").change(function() {
						n._paraPool[8] = $(this).val()
					}), t.find(".boxplot-show-average").change(function() {
						n._paraPool[2] = $(this).prop("checked")
					}), t.find(".boxplot-vertical").change(function() {
						n._paraPool[9] = $(this).prop("checked")
					})
				}, u.prototype.applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t !== undefined && t !== null ? t + "," : ",";
					n = this.removeContinuousComma(n), u = "=BOXPLOTSPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						type: 10,
						dataRange: this._paraPool[0]
					}), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u
			}(u), i.BoxPlotSparklineDialog = y, b = function(n) {
				function u() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.cascadeSparklineDialog.cascadeSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".cascade-sparkline-dialog", r)
				}
				return __extends(u, n), u.prototype._init = function() {
					this._defaultValue = {
						colorPositive: "#8CBF64",
						colorNegative: "#D6604D",
						vertical: !1
					}, this.createColorPicker("cascade-positive-color-frame", "cascade-positive-color-span", "cascade-positive-color-picker"), this.createColorPicker("cascade-negative-color-frame", "cascade-negative-color-span", "cascade-negative-color-picker"), this._attachEvent()
				}, u.prototype.applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t !== undefined && t !== null ? t + "," : ",";
					n = this.removeContinuousComma(n), u = "=CASCADESPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						type: 11,
						dataRange: this._paraPool[0],
						parameterSets: this._paraPool
					}), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, u.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".cascade-points-range").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".cascade-point-index").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".cascade-mini").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".cascade-maxi").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), $(".cascade-positive-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[5] = '"' + r + '"'
					}), $(".cascade-negative-color-picker").bind(this.colorValueChanged, function(t, i) {
						var r = i.value && i.value.color ? i.value.color : "transparent";
						n._paraPool[6] = '"' + r + '"'
					}), t.find(".cascade-labels-range").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".cascade-vertical").change(function() {
						n._paraPool[7] = $(this).prop("checked")
					})
				}, u.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateCascadeSparklineDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex()), this._setPointIndexStatus()
				}, u.prototype._setPointIndexStatus = function() {
					var n = this.sheet.getSelections(),
						t;
					n && n.length > 0 && (t = n[0], n.length > 1 || t.rowCount > 1 || t.colCount > 1 ? this._element.find(".cascade-point-index").attr("disabled", "disabled") : this._element.find(".cascade-point-index").removeAttr("disabled"))
				}, u.prototype._updateCascadeSparklineDialog = function(n, u) {
					var h = i.util.parseFormulaSparkline(n, u),
						f, a, l;
					if (h && h.args) {
						f = h.args;
						if (f && f.length > 0) {
							var s = this._defaultValue,
								y = i.util.unParseFormula(f[0], n, u),
								v = i.util.unParseFormula(f[1], n, u),
								p = i.util.unParseFormula(f[2], n, u),
								b = i.util.unParseFormula(f[3], n, u),
								w = i.util.unParseFormula(f[4], n, u),
								o = r.parseColorExpression(f[5], n, u),
								e = r.parseColorExpression(f[6], n, u),
								c = f[7] instanceof t.Calc.Expressions.BooleanExpression ? f[7].value : null;
							this.updateTextBox("cascade-points-range", y), this.updateTextBox("cascade-point-index", v), this.updateTextBox("cascade-maxi", w), this.updateTextBox("cascade-mini", b), this.updateTextBox("cascade-labels-range", p), this.updateColorPicker("cascade-positive-color-picker", "cascade-positive-color-span", o === null ? s.colorPositive : o), this.updateColorPicker("cascade-negative-color-picker", "cascade-negative-color-span", e === null ? s.colorNegative : e), this.updateCheckBox("cascade-vertical", c === null ? s.vertical : c), a = e ? '"' + e + '"' : null, l = o ? '"' + o + '"' : null, this._paraPool = [y, v, p, b, w, l, a, c]
						} else this._reset()
					}
				}, u.prototype._reset = function() {
					this.updateTextBox("cascade-points-range", ""), this.updateTextBox("cascade-point-index", ""), this.updateTextBox("cascade-maxi", ""), this.updateTextBox("cascade-mini", ""), this.updateTextBox("cascade-labels-range", ""), this.updateColorPicker("cascade-positive-color-picker", "cascade-positive-color-span", this._defaultValue.colorPositive), this.updateColorPicker("cascade-negative-color-picker", "cascade-negative-color-span", this._defaultValue.colorNegative), this.updateCheckBox("cascade-vertical", this._defaultValue.vertical)
				}, u
			}(u), i.CascadeSparklineDialog = b, w = function(n) {
				function r() {
					var t = this,
						r = {
							modal: !0,
							title: i.res.paretoSparklineDialog.paretoSparklineSetting,
							minWidth: 500,
							resizable: !1,
							buttons: [{
								text: i.res.ok,
								click: function() {
									t.applySetting()
								}
							}, {
								text: i.res.cancel,
								click: function() {
									t.close(), i.wrapper.setFocusToSpread()
								}
							}]
						};
					n.call(this, "./dialogs/dialogs2.html", ".pareto-sparkline-dialog", r)
				}
				return __extends(r, n), r.prototype._init = function() {
					this._defaultValue = {
						label: 0,
						vertical: !1
					}, this._attachEvent()
				}, r.prototype.applySetting = function() {
					for (var f = this._paraPool, n = "", t, u, r = 0; r < f.length; r++) t = f[r], n += t !== undefined && t !== null ? t + "," : ",";
					n = this.removeContinuousComma(n), u = "=PARETOSPARKLINE(" + n + ")", i.actions.doAction("setFormulaSparkline", i.wrapper.spread, {
						formula: u,
						type: 12,
						dataRange: this._paraPool[0],
						parameterSets: this._paraPool
					}), this.updateFormulaBar(), this.close(), i.wrapper.setFocusToSpread()
				}, r.prototype._attachEvent = function() {
					var n = this,
						t = n._element;
					t.find(".pareto-points").keyup(function() {
						n._paraPool[0] = $(this).val()
					}), t.find(".pareto-point-index").keyup(function() {
						n._paraPool[1] = $(this).val()
					}), t.find(".pareto-color-range").keyup(function() {
						n._paraPool[2] = $(this).val()
					}), t.find(".pareto-target").keyup(function() {
						n._paraPool[3] = $(this).val()
					}), t.find(".pareto-target2").keyup(function() {
						n._paraPool[4] = $(this).val()
					}), t.find(".pareto-highlightPosition").keyup(function() {
						n._paraPool[5] = $(this).val()
					}), t.find(".pareto-label").change(function() {
						n._paraPool[6] = $(this).val()
					}), t.find(".pareto-vertical").change(function() {
						n._paraPool[7] = $(this).prop("checked")
					})
				}, r.prototype._beforeOpen = function() {
					this._paraPool = [], this.sheet = i.wrapper.spread.getActiveSheet(), this._updateParetoSparklineDialog(this.sheet.getActiveRowIndex(), this.sheet.getActiveColumnIndex()), this._setPointIndexStatus()
				}, r.prototype._setPointIndexStatus = function() {
					var n = this.sheet.getSelections(),
						t;
					n && n.length > 0 && (t = n[0], n.length > 1 || t.rowCount > 1 || t.colCount > 1 ? this._element.find(".pareto-point-index").attr("disabled", "disabled") : this._element.find(".pareto-point-index").removeAttr("disabled"))
				}, r.prototype._updateParetoSparklineDialog = function(n, r) {
					var o = i.util.parseFormulaSparkline(n, r),
						u;
					if (o && o.args) {
						u = o.args;
						if (u && u.length > 0) {
							var l = this._defaultValue,
								a = i.util.unParseFormula(u[0], n, r),
								c = i.util.unParseFormula(u[1], n, r),
								y = i.util.unParseFormula(u[2], n, r),
								v = i.util.unParseFormula(u[3], n, r),
								s = i.util.unParseFormula(u[4], n, r),
								h = i.util.unParseFormula(u[5], n, r),
								e = u[6] instanceof t.Calc.Expressions.DoubleExpression ? u[6].value : null,
								f = u[7] instanceof t.Calc.Expressions.BooleanExpression ? u[7].value : null;
							this.updateTextBox("pareto-points", a), this.updateTextBox("pareto-point-index", c), this.updateTextBox("pareto-color-range", y), this.updateTextBox("pareto-highlightPosition", h), this.updateTextBox("pareto-target", v), this.updateTextBox("pareto-target2", s), this.updateSelect("pareto-label", e === null ? l.label : e), this.updateCheckBox("cascade-vertical", f === null ? l.vertical : f), this._paraPool = [a, c, y, v, s, h, e, f]
						} else this._reset()
					}
				}, r.prototype._reset = function() {
					this.updateTextBox("pareto-points", ""), this.updateTextBox("pareto-point-index", ""), this.updateTextBox("pareto-color-range", ""), this.updateTextBox("pareto-highlightPosition", ""), this.updateTextBox("pareto-target", ""), this.updateTextBox("pareto-target2", ""), this.updateSelect("pareto-label", this._defaultValue.label), this.updateCheckBox("cascade-vertical", this._defaultValue.vertical)
				}, r
			}(u), i.ParetoSparklineDialog = w, r = function() {
				function n() {}
				return n.parseColorExpression = function(n, r, u) {
					var e, f;
					if (!n) return null;
					e = i.wrapper.spread.getActiveSheet();
					if (n instanceof t.Calc.Expressions.StringExpression) return n.value;
					else if (n instanceof t.Calc.Expressions.MissingArgumentExpression) return null;
					else {
						f = null;
						try {
							f = i.util.unParseFormula(n, r, u)
						} catch (o) {}
						return t.Calc.evaluateFormula(e, f, r, u)
					}
				}, n
			}()
		})(t.designer || (t.designer = {}));
		var i = t.designer
	})(n.spread || (n.spread = {}));
	var t = n.spread
})(sku || (sku = {}))
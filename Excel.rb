# -*- coding: utf-8 -*-
#
# excel.rb
#

require File.expand_path( File.dirname(__FILE__) + '/win32ole-ext.rb' )
require File.expand_path( File.dirname(__FILE__) + '/Util.rb' )

module Excel

	#----------------------------------------------
	# @biref	Excel オブジェクトを生成する
	# @param	visible			false でバックグラウンドでexcel起動
	# @param	display_alerts	false で特定の警告やメッセージを表示しない
	# @return	Excel オブジェクト
	#----------------------------------------------
	def Excel.new(visible = false, display_alerts = false)
		excel = WIN32OLE.new_with_const('Excel.Application', Excel)
		excel.visible = visible
		excel.displayAlerts = display_alerts
		excel.screenUpdating = visible					# 画面更新表示/非表示(visibleと合わせて設定する)
#		excel.calculation = Excel::XlCalculationManual	# 再計算を手動でやる（自動の再計算を止める）
		return excel
	end

	#----------------------------------------------
	# @biref	Excel の起動と終了のEAM（Execute Around Method）
	# @param	visible			false でバックグラウンドでexcel起動
	# @param	displayAlerts	false で特定の警告やメッセージを表示しない
	# @param	block			ブロック引数
	# @note		ランタイムエラーが起こった場合Excelを終了します
	#----------------------------------------------
	def Excel.runDuring(visible = false, display_alerts = false, &block)
		begin
			excel = new(visible, display_alerts)
			block.call(excel)
		ensure
			if( excel != nil )
				excel.Quit
			end
		end
	end

	#----------------------------------------------
	# @biref	指定したワークブックを開く
	#----------------------------------------------
	def Excel.openWb( excel, file_path, pass = nil )

		file_path	= File.expand_path( file_path )
		file_path	= file_path.gsub( "\\", "/" )
		if( pass == nil )
			return ( excel.workbooks.open( {'filename'=> file_path, 'updatelinks'=> 0} ) )
		else
			return ( excel.workbooks.open( {'filename'=> file_path, 'updatelinks'=> 0, 'password'=>"#{pass}"}) )
		end
	end

	#----------------------------------------------
	# @biref	新規ワークブックを生成
	#----------------------------------------------
	def Excel.createWb( excel )
		return excel.workbooks.add()
	end

	#----------------------------------------------
	# @biref	指定したワークブックを保存、終了
	#----------------------------------------------
	def Excel.saveAndClose( wb, file_path = nil )
		if ( file_path != nil )
			fso = WIN32OLE.new('Scripting.FileSystemObject')
			wb.saveAs( fso.GetAbsolutePathName("#{file_path}") )
		else
			wb.save()
		end
		wb.close(0)
	end

	#----------------------------------------------
	# @biref	指定した名前のシートがあるか
	#----------------------------------------------
	def Excel.existSheet( wb, sheet_name )

        # シート名のチェック
        sheet_name = sheet_name.encode( Encoding::WINDOWS_31J )
		is_exist_sheet = false
		wb.worksheets.each { |ws|
			if( ws.name == sheet_name )
				return true
			end
		}

		print_wb_name	 = wb.name.encode( Encoding::UTF_8 )
		print_sheet_name = sheet_name.encode( Encoding::UTF_8 )
		error_str	=	"「#{print_wb_name}」に「#{print_sheet_name}」シートがありません。\n"
		error_str	+= "シートが存在するか\n"
		error_str	+= "シート名が「#{print_sheet_name}」になっているかお確かめ下さい"
		assertLogPrintFalse( error_str )
		return false
	end

	#----------------------------------------------
	# @biref	指定した文字列の列番号を返す
	# @param	ws			ワークシート
	# @param	search_str	チェックするフィールド名
	# @param	search_row	フィールド（は1行目のはず）
	# @return	列番号
	#----------------------------------------------
	def Excel.getColumn(ws, search_str = "", search_row = 1)

		# 文字列の検索
		search_result = ws.Rows(search_row).Find('What' => search_str)

		# 検索結果
		if (search_result == nil) then
			utf_search_str	= search_str.encode( Encoding::UTF_8 )
			error_str		= "Not Found Column Name !! 『#{utf_search_str}』"
			assertLogPrintFalse( error_str )
			return 0
		else
			return search_result.Column
		end
	end

	#----------------------------------------------
	# @biref	指定した文字列の行番号を返す
	# @param	ws				ワークシート
	# @param	search_str		チェックするフィールド名
	# @param	search_column	検索する列番号
	# @return	列番号
	#----------------------------------------------
	def Excel.getRow(ws, search_str, search_column)

		# 文字列の検索
		search_result = ws.Columns(search_column).Find('What' => search_str)

		# 検索結果
		if (search_result == nil) then
			utf_search_str	= search_str.encode( Encoding::UTF_8 )
			error_str		= "Not Found Row Name !! 『#{utf_search_str}』"
			assertLogPrintFalse( error_str )
			return 0
		else
			return search_result.Row
		end
	end

	#----------------------------------------------
	# @biref	Range の文字列を算出する（開始セルと行数から）
	# @param	range_st_column	Range 開始列名
	# @param	range_st_row	Range 開始行名
	# @param	row_count		行数
	#----------------------------------------------
	def Excel.calcRangeStr( range_st_column, range_st_row, row_count )

		range_str = "#{range_st_column}#{range_st_row}:"
		range_str += "#{range_st_column}#{row_count-1}"
		return range_str
	end

	#----------------------------------------------
	# @biref	指定したセルのデータを返す
	# @param	ws				ワークシート
	# @param	row_index		行番号
	# @param	column_index	列番号
	# @return	データ
	#----------------------------------------------
	def Excel.getCellValue(ws, row_index = 1, column_index = 1)
		return ws.Cells.Item(row_index, column_index).Value
	end

	#----------------------------------------------
	# @biref	指定したセルのデータを返す（列名指定ver）
	# @param	ws				ワークシート
	# @param	row_index		行番号
	# @param	column_name		列名(1行目前提)
	# @return	データ
	#----------------------------------------------
	def Excel.getCellValueWithColumnName(ws, row_index, column_name)

		column_index = getColumn( ws, "#{column_name}" )
		return ws.Cells.Item(row_index, column_index).Value
	end

	#----------------------------------------------
	# @biref	指定したセルのデータを設定する
	# @param	ws				ワークシート
	# @param	row_index		行番号
	# @param	column_index	列番号
	# @param	value			設定する値
	#----------------------------------------------
	def Excel.setCellValue(ws, row_index, column_index, value)
		ws.Cells.Item(row_index, column_index).Value = value
	end

	#----------------------------------------------
	# @biref	Excel シートコピー
	# @param	src_wb			コピー元のワークブック
	# @param	src_ws_name		コピー元のワークシートネーム
	# @param	dst_wb			コピー先のワークブック
	# @param	dst_ws_number	コピー先のワークシート番号
	#----------------------------------------------
	def Excel.sheetCopy( src_wb, src_ws_name, dst_wb, dst_ws_number )

		# シートをコピー
		if( existSheet( src_wb, src_ws_name ) == true )
			ws_temp	= src_wb.worksheets( "#{src_ws_name}" )
			ws_temp.copy( {'After'=> dst_wb.worksheets(dst_ws_number)} )
		end
	end

	#----------------------------------------------
	# @biref	Excel シートコピー(シート番号指定版)
	# @param	src_wb			コピー元のワークブック
	# @param	src_ws_number	コピー元のワークシート番号
	# @param	dst_wb			コピー先のワークブック
	# @param	dst_ws_number	コピー先のワークシート番号
	#----------------------------------------------
	def Excel.sheetCopyNumber( src_wb, src_ws_number, dst_wb, dst_ws_number )

		ws_temp	= src_wb.worksheets( src_ws_number )
		ws_temp.copy( {'After'=> dst_wb.worksheets(dst_ws_number)} )
	end

	#----------------------------------------------
	# @biref	Excel 行コピー＆ペースト(挿入)
	# @param	src_ws		コピー元のワークシート
	# @param	src_row		コピー元の行番号
	# @param	dst_ws		コピー先のワークシート
	# @param	dst_row		コピー先の行番号
	#----------------------------------------------
	def Excel.rowCopyAndInsert( src_ws, src_row, dst_ws, dst_row )

		src_ws.range("#{src_row}:#{src_row}").copy
		dst_ws.range("#{dst_row}:#{dst_row}").insert
		dst_ws.range("#{dst_row}:#{dst_row}").pastespecial
	end

	#----------------------------------------------
	# @biref	Excel 範囲コピー
	# @param	src_ws		コピー元のワークシート
	# @param	src_range	コピー元の範囲指定
	# @param	dst_ws		コピー先のワークシート
	# @param	dst_range	コピー先の範囲指定
	# @note 	普通に Value を設定するより高速
	#----------------------------------------------
	def Excel.rangeCopy( src_ws, src_range, dst_ws, dst_range )
		src_ws.range( src_range ).copy( {'Destination'=> dst_ws.range( dst_range )} )
	end

	#----------------------------------------------
	# @biref	指定文字列に色をつける
	# @param	ws				ワークシート
	# @param	cell_row		セルの行番号
	# @param	cell_column		セルの列番号
	# @param	src_str			文字列
	# @param	color_str		src_strの中で色をつけたい文字列
	#----------------------------------------------
	def Excel.setStringColor( ws, cell_row, cell_column, src_str, color_str )

		# 文字色を赤色にする
		prefix_str_index = src_str.index( color_str ) + 1
		color_str_length = color_str.length
		ws.Cells.Item( cell_row, cell_column ).Characters( {'Start' => prefix_str_index, 'Length' => color_str_length}).Font.ColorIndex = 3
	end

	#----------------------------------------------
	# @biref	指定列を表示／非表示
	# @param	ws			ワークシート
	# @param	column		列番号
	# @param	is_visible	表示にするか
	#----------------------------------------------
	def Excel.setVisibleColumns( ws, column, is_visible )
		ws.Cells.Columns(column).Hidden	= !is_visible
	end

	#----------------------------------------------
	# @biref	指定列を表示／非表示
	# @param	ws			ワークシート
	# @param	row			行番号
	# @param	is_visible	表示にするか
	#----------------------------------------------
	def Excel.setVisibleRows( ws, row, is_visible )
		ws.Cells.Rows.(row).Hidden = !is_hidden
	end

	#----------------------------------------------
	# @biref	指定セルのロック設定／解除
	# @param	ws			ワークシート
	# @param	cell_row	セルの行番号
	# @param	cell_column	セルの列番号
	# @param	is_lock		ロックするか
	#----------------------------------------------
	def Excel.setLockCell( ws, cell_row, cell_column, is_lock )
		ws.Cells.Item(cell_row, cell_column).Locked = is_lock
	end

	#----------------------------------------------
	# @biref	指定セルにコメントを追加する
	# @param	ws			ワークシート
	# @param	cell_row	セルの行番号
	# @param	cell_column	セルの列番号
	# @param	comment		コメント
	#----------------------------------------------
	def Excel.setAddCommentCell( ws, cell_row, cell_column, comment )
		ws.Cells.Item(cell_row, cell_column).AddComment( comment )
	end

	#----------------------------------------------
	# @biref	シートを表示／非表示
	# @param	ws			ワークシート
	# @param	is_visible	表示にするか
	#----------------------------------------------
	def Excel.setVisibleSheet( ws, is_visible )
		ws.visible = is_visible
	end

	#----------------------------------------------
	# @biref	シート保護の設定／解除
	# @param	ws			ワークシート
	# @param	is_protect	保護するか
	#----------------------------------------------
	def Excel.setProtectSheet( ws, is_protect )
		if ( is_protect == true )
			ws.Protect
		else
			ws.UnProtect
		end
	end

	#----------------------------------------------
	# @biref	指定ワークシートの色を赤に
	# @param	ws		ワークシート
	#----------------------------------------------
	def Excel.setSheetRedColor( ws )
		ws.Tab.ColorIndex = 3
	end

	#----------------------------------------------
	# @biref	指定年月日が土日の際にシート色をつける
	# @param	ws		ワークシート
	# @param	w_day	曜日
	#----------------------------------------------
	def Excel.setSheetColorWithWeekend( ws, year, month, day )

		# 日曜
		w_day = calcWeekDay( year, month, day )
		if( w_day == "日" )
			ws.Tab.ColorIndex = 3
		# 土曜
		elsif( w_day == "土" )
			ws.Tab.ColorIndex = 5
		# それ以外
		else
			ws.Tab.ColorIndex = Excel::XlNone
		end
	end

	#----------------------------------------------
	# @biref	指定ワークシートのシート色が赤か
	# @param	ws		ワークシート
	#----------------------------------------------
	def Excel.isWsColorRed( ws )
		if( ws.Tab.ColorIndex == 3 )
			return true
		else
			return false
		end
	end

	#----------------------------------------------
	# @biref	Excel のデフォルトのシートを削除
	# @param	wb		ワークブック
	#----------------------------------------------
	def Excel.deleteDefaultSheet( wb )

		# [Sheet1][Sheet2][Sheet3] を削除
		(1..3).each{|num|
			ws = wb.worksheets("Sheet#{num}")
			if( ws != nil)
				ws.delete()
			end
		}
	end

	#----------------------------------------------
	# @biref	指定ワークブックのシート数を取得
	# @param	wb		ワークブック
	#----------------------------------------------
	def Excel.getSheetCount( wb )
		return wb.worksheets.count()
	end

	#----------------------------------------------
	# @biref	アクティブウィンドウのスクロール設定
	# @param    scroll_row		スクロールさせたい行
	# @param    scroll_column	スクロールさせたい列
	#----------------------------------------------
	def Excel.setScrollWithActiveWindow( excel, scroll_row, scroll_column )
		excel.ActiveWindow.ScrollRow = scroll_row
		excel.ActiveWindow.ScrollColumn = scroll_column
	end

	#----------------------------------------------
	# @biref	シート全体を自動調整
	# @param    ws  自動調整したいワークシート
	#----------------------------------------------
	def Excel.autoFitCellAll( ws )

		# シートをアクティブに
		ws.Activate

		# シート全体を選択
		ws.Cells.Select

		# 行／列を自動調整
		ws.Cells.Rows.AutoFit
		ws.Cells.Columns.AutoFit
	end

	#----------------------------------------------
	# @biref	指定した値を検索して上書きする
	# @param	ws	操作するワークシート
	# @param	serach_str	検索したい値
	# @param	set_value	セットしたい値
	# @return	列番号
	#----------------------------------------------
	def Excel.resetData( ws, search_str, set_value )
		p "resetData()"

		# 検索
		found_cell = ws.Cells.Find('What' => search_str)

		# 最初のセルを記憶
		if found_cell == nil
			assertLogPrintFalse( "not found rewright data..." )
			return
		else
			first_cell = found_cell
		end

		# 最初のセルになるまでループ
		cellList = []
		begin
			# セルをリストへ
			cellList.push(found_cell)

			# 見つかったセルの次のセルを検索。最終的には最初に戻ってくる
			found_cell = ws.Cells.FindNext(found_cell)

		end while (found_cell.Address != first_cell.Address)
		p "found_cell => #{cellList.size}"

		# 検索にHITしたセルに指定された値を設定
		cellList.each { |cell|
			ws.Cells.Item(cell.Row, cell.Column).Value = set_value
		}
	end

end

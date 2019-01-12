# -*- coding: utf-8 -*-
#
#　excel.rb
#

require File.expand_path( File.dirname(__FILE__) + '/win32ole-ext.rb' )

module Excel

#	excel.ActiveCell.Offset(0, -1).Activate

	#----------------------------------------------
	# @biref	Excel オブジェクトを生成する
	# @parm		visible			false でバックグラウンドでexcel起動
	# @parm		displayAlerts	false で特定の警告やメッセージを表示しない
	# @return	Excel オブジェクト
	#----------------------------------------------
	def Excel.new(visible = true, displayAlerts = false)
		excel = WIN32OLE.new_with_const('Excel.Application', Excel)
		excel.visible = visible
		excel.displayAlerts = displayAlerts 
		excel.screenUpdating = false
#		excel.calculation = Excel::XlCalculationManual
		return excel 
	end

	#----------------------------------------------
	# @biref	Excel の起動と終了のEAM 
	# @parm		visible			false でバックグラウンドでexcel起動
	# @parm		displayAlerts	false で特定の警告やメッセージを表示しない
	# @parm		block			ブロック引数
	# @note		ランタイムエラーが起こった場合Excel終了します
	#----------------------------------------------
	def Excel.runDuring(visible = true, displayAlerts = false, &block) 
		begin
			excel = new(visible, displayAlerts) 
			block.call(excel)
		ensure
#			excel.visible = true	
#			excel.displayAlerts = true
#			excel.screenUpdating = true
#			excel.Calculation = Excel::XlCalculationAutomatic
			excel.Quit 
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
	# @biref	指定した名前のシートがあるか
	#----------------------------------------------
	def Excel.existSheet( wb, sheet_name )

        # シート名のチェック
        sheet_name = sheet_name.encode( "Windows-31J" )        
		is_exist_sheet = false
		wb.worksheets.each { |ws|

			if( ws.name == sheet_name )
				return true
			end		
		}

		print_wb_name	 = wb.name.encode( "UTF-8" )		
		print_sheet_name = sheet_name.encode( "UTF-8" )
		error_str	=	"「#{print_wb_name}」に「#{print_sheet_name}」シートがありません。\n"
		error_str	+= "シートが存在するか\n"
		error_str	+= "シート名が「#{print_sheet_name}」になっているかお確かめ下さい"
		assertLogPrintFalse( error_str )
		return false
	end
	
	#----------------------------------------------
	# @biref	指定した文字列の列番号を返す
	# @parm		ws			ワークシート
	# @parm		search_str	チェックするフィールド名
	# @parm		search_row	フィールドは1行目
	# @return	列番号
	#----------------------------------------------
	def Excel.getColumn(ws, search_str ="", search_row = 1)
	#	p "getColumn()"

		#文字列の検索
		search_result = ws.Rows(search_row).Find('What' => search_str)
		
		#検索結果
		if (search_result == nil) then
#			p "Not FoundName !!" + "『" + search_str + "』"
			p "Not FoundName !!"
			return 0
		else
			return search_result.Column
		end
	end

	#----------------------------------------------
	# @biref	指定した文字列の行番号を返す
	# @parm		ws				ワークシート
	# @parm		search_column	検索する行番号
	# @parm		search_str		チェックするフィールド名
	# @return	列番号
	#----------------------------------------------
	def Excel.getRow(ws, search_column, search_str )
	#	p "getRow()"

		#文字列の検索
		search_result = ws.Columns(search_column).Find('What' => search_str)
		
		#検索結果
		if (search_result == nil) then
#			p "Not FoundName !!" + "『" + search_str + "』"
			p "Not FoundName !!"
			return 0
		else
			return search_result.Row
		end
	end

	#----------------------------------------------
	# @biref	Range　の文字列を算出する（開始セルと行数から）
	# @parm		range_st_colmn	Range　開始列名
	# @parm		range_st_row	Range　開始行名
	# @parm		row_count		行数
	#----------------------------------------------
	def Excel.calcRangeStr( range_st_colmn, range_st_row, row_count )

		range_str = "#{range_st_colmn}#{range_st_row}:"
		range_str += "#{range_st_colmn}#{row_count-1}"
		
		return range_str
	end
	
	#----------------------------------------------
	# @biref	指定したパラメータのデータを返す
	# @parm		ws			ワークシート
	# @parm		row_index	行番号
	# @parm		param_name	パラメータ名(1行目前提)
	# @return	データ
	#----------------------------------------------
	def Excel.getParamValue(ws, row_index =1, param_name ="")

		clumn_index = getColumn( ws, "#{param_name}" )
		return ws.Cells.Item(row_index, clumn_index).Value
	end

	#----------------------------------------------
	# @biref	指定したセルのデータを返す
	# @parm		ws			ワークシート
	# @parm		row_index	行番号
	# @parm		clumn_index	列番号
	# @return	データ
	#----------------------------------------------
	def Excel.getCellValue(ws, row_index =1, clumn_index = 1)
	#	p "getCellValue()"

		return ws.Cells.Item(row_index, clumn_index).Value
	end

	#----------------------------------------------
	# @biref	Excel シートコピー
	# @parm		src_wb			コピー元のワークブック
	# @parm		src_ws_name		コピー元のワークシートネーム
	# @parm		dst_wb			コピー先のワークブック
	# @parm		dst_ws_number	コピー先のワークシート番号
	#----------------------------------------------
	def Excel.sheetCopy( src_wb, src_ws_name, dst_wb, dst_ws_number )

		if( existSheet( src_wb, src_ws_name ) == true )

			# シートをコピー
			ws_temp	= src_wb.worksheets( "#{src_ws_name}" )
			ws_temp.copy( {'After'=> dst_wb.worksheets(dst_ws_number)} )
		end
	end

	#----------------------------------------------
	# @biref	Excel シートコピー(シート番号指定版)
	# @parm		src_wb			コピー元のワークブック
	# @parm		src_ws_number	コピー元のワークシート番号
	# @parm		dst_wb			コピー先のワークブック
	# @parm		dst_ws_number	コピー先のワークシート番号
	#----------------------------------------------
	def Excel.sheetCopyNumber( src_wb, src_ws_number, dst_wb, dst_ws_number )

		ws_temp	= src_wb.worksheets( src_ws_number )
		ws_temp.copy( {'After'=> dst_wb.worksheets(dst_ws_number)} )
	end

	#----------------------------------------------
	# @biref	Excel 行コピー＆ペースト(挿入)
	# @parm		src_ws		コピー元のワークシート
	# @parm		src_row		コピー元の行番号
	# @parm		dst_ws		コピー先のワークシート
	# @parm		dst_row		コピー先の行番号
	#----------------------------------------------
	def Excel.rowCopyAndInsert( src_ws, src_row, dst_ws, dst_row )

		src_ws.range("#{src_row}:#{src_row}").copy
		dst_ws.range("#{dst_row}:#{dst_row}").insert
		dst_ws.range("#{dst_row}:#{dst_row}").pastespecial
	end

	#----------------------------------------------
	# @biref	Excel 範囲コピー（高速版）
	# @parm		src_ws		コピー元のワークシート
	# @parm		src_range	コピー元の範囲指定
	# @parm		dst_ws		コピー先のワークシート
	# @parm		dst_range	コピー先の範囲指定
	#----------------------------------------------
	def Excel.rangeCopyFast( src_ws, src_range, dst_ws, dst_range )
		src_ws.range( src_range ).copy( {'Destination'=> dst_ws.range( dst_range )} )
	end

	#----------------------------------------------
	# @biref	Excel 範囲コピー
	# @parm		src_ws		コピー元のワークシート
	# @parm		src_range	コピー元の範囲指定
	# @parm		dst_ws		コピー先のワークシート
	# @parm		dst_range	コピー先の範囲指定
	#----------------------------------------------
	def Excel.rangeCopy( src_ws, src_range, dst_ws, dst_range )

		dst_ws.range( "#{dst_range}" ).Value = src_ws.range( "#{src_range}" ).Value
	end	

	#----------------------------------------------
	# @biref	指定文字列に色をつける
	# @parm		ws				ワークシート
	# @parm		cell_row		セルの行番号
	# @parm		cell_clumn		セルの列番号
	# @parm		src_str			文字列
	# @parm		color_str		src_strの中で色をつけたい文字列
	#----------------------------------------------
	def Excel.setStringColor( ws, cell_row, cell_clumn, src_str, color_str )
	
		# 文字色を赤色にする
		prefix_str_index = src_str.index( color_str ) + 1
		color_str_length = color_str.length
		ws.Cells.Item( cell_row, cell_clumn ).Characters( {'Start' => prefix_str_index, 'Length' => color_str_length}).Font.ColorIndex = 3
	end

	
	#----------------------------------------------
	# @biref	Excel のデフォルトのシートを削除
	# @parm		wb		ワークブック
	#----------------------------------------------
	def Excel.deleteDefaultSheet( wb )
		(1..3).each{|num|
			ws = wb.worksheets("Sheet#{num}")
			if( ws != nil)
				ws.delete()			
			end
		}
	end

	#----------------------------------------------
	# @biref	シート全体を自動調整
	# @param    ws  自動調整したいワークシート
	# @return   NULL
	#----------------------------------------------
	def Excel.autoFitDataRange(ws)

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
	# @param	set_str		セットしたい値
	# @return	列番号
	#----------------------------------------------
	def Excel.resetData(ws, search_str, set_str)
		p "resetData()"
		
		# 検索
		foundCell = ws.Cells.Find('What' => search_str)
	#                        What:=search_str, _
	#                        LookIn:=xlValues, _
	#                        LookAt:=xlWhole, _
	#                        SearchOrder:=xlByRows, _
	#                        MatchCase:=False, _
	#                        MatchByte:=False)
	
		# 最初のセルを記憶
		if foundCell == nil
			p "not found rewright data...."
			return
		else
			firstCell = foundCell
		end

		cellList = []
		
		# 最初のセルになるまでループ
		begin
			# セルをリストへ
			cellList.push(foundCell)			

			# 見つかったセルの次のセルを検索。最終的には最初に戻ってくる
			foundCell = ws.Cells.FindNext(foundCell)
		
		end while (foundCell.Address != firstCell.Address)

		# 最初のセルになるまでループ
		p "foundCell => #{cellList.size}"

		# 検索にHITしたセルを選択
		cellList.each { |cell|
			ws.Cells.Item(cell.Row, cell.Column).Value = set_str
		}
	end
	   
end
 
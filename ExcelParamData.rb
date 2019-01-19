# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )

# ==========================="
# src
# ==========================="
class ExcelParamData
	public
	def initialize(wb_path, ws_name, param_name_hash)
		@wb_path = wb_path
		@ws_name = ws_name
		@param_name_hash = param_name_hash
		@param_list = Array.new
		@param_list.clear

		assertLogPrintNotFoundFile( @wb_path )
		setData()
	end

	# パラメータのリストを取得
	def getParamList()
		return @param_list
	end

	private
	def setData()

		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param = Excel.openWb( excel, @wb_path )
			ws_param = wb_param.worksheets( @ws_name )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do

				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				# パラメータを取得してpush
				is_next = false
				param = Hash.new
				@param_name_hash.each  { |key, value|
					column_index = Excel.getColumn(ws_param, "#{value}")
					cell_value = Excel.getCellValue(ws_param, recode.row, "#{column_index}".to_i)
					# 空白行はスキップ
					if (cell_value == "" or cell_value == nil)
						is_next = true
						break
					end
					param[ :"#{key}" ] = cell_value
				}
				next if (is_next == true)
				@param_list.push( param )
			end
			wb_param.close(0)
		end
		errorCheck()
	end

	def errorCheck()
		@param_list.each { |param|
			@param_name_hash.each  { |key, value|
				data = param[ :"#{key}" ]
				if( data == "" or data == nil )
					error_str = "Parameter Error!!"
					error_str = error_str + "「#{value}」が未入力です。"
					assertLogPrintFalse( "#{error_str}" )
				end
			}
		}
	end

end

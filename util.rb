# -*- coding: utf-8 -*-

require 'benchmark'

IS_DEBUG	= true	# デバッグ機能を利用する際に使用する

# 月ごとの日数
MONTH_DAYS = [ 31, 28, 31, 30, 31, 30,
						  31, 31, 30, 31, 30, 31 ] 

if( IS_DEBUG == true )
	def dbgPrint( *args )
		p( args )
	end

	def dbgPuts( *args )
		puts( args )
	end
else # IS_DEBUG
	def dbgPrint( *args )
	end

	def dbgPuts( *args )
	end
end # IS_DEBUG

def releasePuts( args, file = nil )
	puts( args.encode( Encoding::UTF_8 )  )
	if( file != nil )
		file.puts( args.encode( Encoding::UTF_8 ) )
	end
end

def releasePrint( args, file = nil )
	p( args.encode( Encoding::UTF_8 )  )
	if( file != nil )
		file.print( args.encode( Encoding::UTF_8 ) )
	end
end

# --------------------------------------------
# コンソールの色取得
# --------------------------------------------
def getConsoleColor( color_name )

	if( "#{color_name}" == "black" ) 	#
		return "0"
	elsif( "#{color_name}" == "navy" ) 	# 暗い青
		return "1"
	elsif( "#{color_name}" == "green" )
		return "2"
	elsif( "#{color_name}" == "teal" ) 	# 青緑
		return "3"
	elsif( "#{color_name}" == "maroon" )# 暗い赤
		return "4"
	elsif( "#{color_name}" == "purple" )
		return "5"
	elsif( "#{color_name}" == "olive" )	# 暗い黄色
		return "6"
	elsif( "#{color_name}" == "silver" )
		return "7"
	elsif( "#{color_name}" == "gray" )
		return "8"
	elsif( "#{color_name}" == "blue" )
		return "9"
	elsif( "#{color_name}" == "lime" )	# 明るい緑
		return "A"
	elsif( "#{color_name}" == "aqua" )	# 水色
		return "B"
	elsif( "#{color_name}" == "red" )
		return "C"
	elsif( "#{color_name}" == "magenta" )# 明るい紫
		return "D"
	elsif( "#{color_name}" == "yellow" )
		return "E"
	elsif( "#{color_name}" == "white" )
		return "F"
	else
		assertLogPrintFalse( "色番号を間違えています => #{color_name}" )
	end
end

# --------------------------------------------
# コンソールの背景色設定
# --------------------------------------------
def setConsoleColor( color_name_bg, color_name_text )
	color_bg	= getConsoleColor( color_name_bg )
	color_text	= getConsoleColor( color_name_text )
	system("color #{color_bg}#{color_text}")
	
	if( color_bg == color_text )
		assertLogPrintFalse( "背景色と文字色に同じ色を指定しています" )
	end
end

# --------------------------------------------
# アサートログ出力
# --------------------------------------------
def assertLogPrintFalse( error_str )
	assertLogPrint( false, error_str )
end

def assertLogPrint( eq, error_str )

	return if ( eq )

	releasePuts ""
	releasePuts "************************** error **************************"
	releasePuts "#{error_str}"
	releasePuts "***********************************************************"
	setConsoleColor( "maroon", "white" )
	raise( "assert" )
end

def assertLogPrintNotFoundFile( file_path )

		file_path = File.expand_path( file_path )
		if( File.exist?( file_path ) == false )
			error_str = "#{File.basename( file_path )} がありません\n"
			error_str += "パス詳細 => #{file_path}"
			assertLogPrintFalse( "#{error_str}" )
		end
end

# --------------------------------------------
# ワーニングログ出力
# --------------------------------------------
def warningLogPrint( warning_str )

	releasePuts "******************** warning ********************"
	releasePuts "#{warning_str}"
	releasePuts "*************************************************"
	releasePuts ""
end

# --------------------------------------------
# 正常終了時のログ出力
# --------------------------------------------
def successLogPrint()

	releasePuts "***********************************************************"
	releasePuts "正常に終了しました"
	releasePuts "***********************************************************"
	setConsoleColor( "navy", "white" )
end

# --------------------------------------------
# アサートログ出力
# --------------------------------------------
def assertLogPrintFalseNotFound( file_path )

	error_str = "#{File.basename( file_path )} がありません"
	assertLogPrintFalse( "#{error_str}" )
end

# --------------------------------------------
# 処理計測
# ブロックを渡して、ベンチマークで計測。処理時間をプリントします
# --------------------------------------------
def calcProcessTime( &block )

	result = Benchmark.realtime { |process|
		block.call
	}
	puts "処理時間 => [#{result}] sec"
end

# --------------------------------------------
# 検索パターン配列の取得
# --------------------------------------------
def getSearchPatternList( root_dir, pattern_array )

	search_list = Array.new()
	pattern_array.each { |pat|
		serach_pat = "#{root_dir}" + "/**/" + "#{pat}"	
		search_list.push( serach_pat )
	}
	
	return search_list
end

#----------------------------------------------
# パターンにマッチするファイルの検索
# @parm		root_dir		検索のルートパス
# @parm		search_pattern	検索パターン
# @return	ファイルリスト
#----------------------------------------------
def getSearchFile( root_dir, search_pattern )

	# パターンにマッチするファイルパスを追加
	file_list = Array.new
	file_list.clear
	
	serach_pat_list = getSearchPatternList( "#{root_dir}", search_pattern )
	Dir.glob( serach_pat_list ) do |file_path|
		file_list.push( file_path )
	end

	# ascii順に並び替え
	file_list.sort!	
	
	return file_list
end

# --------------------------------------------
# 出力フォルダのファイルを削除
# @parm		root_dir		検索のルートパス
# @parm		search_pattern	検索パターン
# --------------------------------------------
def allClearFile( root_dir, search_pattern )

	# ファイルを削除
	serach_pat_list = getSearchPatternList( "#{root_dir}", search_pattern )	
	Dir.glob( serach_pat_list ) do |file_path|
		FileUtils.rm_r( Dir.glob( "#{file_path}" ) )
	end

	# out フォルダ以下のフォルダを削除
	Dir.glob( "#{root_dir}/**" ) do |file_path|
		FileUtils.rm_r( Dir.glob( "#{file_path}" ) )
	end
end

# --------------------------------------------
# 指定月の日数を返す
# @parm		root_dir		検索のルートパス
# @parm		search_pattern	検索パターン
# --------------------------------------------
def getMonthlyDayCount( month_num )
							
	# 指定月の日数を設定
	monthly_days = MONTH_DAYS[ month_num - 1 ] # [0始まり] と [1始まり] の帳尻合わせ 

	# 閏年を考慮
	if( month_num == 2 && Date.new( year ).leap? )
		monthly_days += 1
	end
	
	return monthly_days
end

#----------------------------------------------
# 数値（年+月） を 「年」「月」に分けて返す（文字列）
#----------------------------------------------
def getSplitCalendar( value )

	calendar = value.to_s	
	return calendar.unpack("a4a2")	# 4文字 / 2文字に分割
end

#----------------------------------------------
# 西暦を平成に変換する
#----------------------------------------------
def getYearNumber( year )

	# year を 1文字 / 3文字に分割 => 下三桁に12を加算
	year_number_array = year.unpack("a1a3")
	return ( year_number_array[1].to_i + 12 )
end

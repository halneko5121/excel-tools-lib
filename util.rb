# -*- coding: utf-8 -*-

require 'benchmark'
require "date"
require "fileutils"

IS_DEBUG = true	# デバッグ機能を利用する際に使用する

# 月ごとの日数
MONTH_DAYS = [ 31, 28, 31, 30, 31, 30,
			   31, 31, 30, 31, 30, 31 ]
# 曜日
WDAYS = ["日", "月", "火", "水", "木", "金", "土"]

if( IS_DEBUG == true )
	def dbgPrint( *args )
		p( args )
	end

	def dbgPuts( *args )
		puts( args )
	end
else
	def dbgPrint( *args )
	end

	def dbgPuts( *args )
	end
end

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
# @param	root_dir		検索のルートパス
# @param	pattern_array	検索パターン配列
# @return	ファイルリスト
#----------------------------------------------
def getSearchFileList( root_dir, pattern_array )

	# パターンにマッチするファイルパスを追加
	file_list = Array.new
	file_list.clear

	search_pat_list = getSearchPatternList( "#{root_dir}", pattern_array )
	Dir.glob( search_pat_list ) do |file_path|
		file_list.push( file_path )
	end

	# ascii順に並び替え
	file_list.sort!

	return file_list
end

# --------------------------------------------
# 出力フォルダのファイルを削除
# @param	root_dir		検索のルートパス
# @param	pattern_array	検索パターン配列
# --------------------------------------------
def allClearFile( root_dir, pattern_array )

	# ファイルを削除
	search_pat_list = getSearchPatternList( "#{root_dir}", pattern_array )
	Dir.glob( search_pat_list ) do |file_path|
		FileUtils.rm_r( Dir.glob( "#{file_path}" ) )
	end

	# out フォルダ以下のフォルダを削除
	Dir.glob( "#{root_dir}/**" ) do |file_path|
		FileUtils.rm_r( Dir.glob( "#{file_path}" ) )
	end
end

#----------------------------------------------
# 指定年月がうるう年か
# @param	year			年
# @param	month_num		月（1 ~ 12）
#----------------------------------------------
def isLeapYear( year, month_num )
	if( month_num == 2 && Date.new( year ).leap? )
		return true
	end

	return false
end

# --------------------------------------------
# 指定年月の日数を返す
# @param	year			年
# @param	month_num		月（1 ~ 12）
# --------------------------------------------
def getMonthlyDayCount( year, month_num )

	# 指定月の日数を設定
	monthly_days = MONTH_DAYS[ month_num - 1 ] # [0始まり] と [1始まり] の帳尻合わせ

	# 閏年を考慮
	if( isLeapYear( year, month_num ) )
		monthly_days += 1
	end

	return monthly_days
end

#----------------------------------------------
# 数値（年+月） を 「年」「月」に分けて返す（文字列）
# 例:201801 => [2018][01]
#----------------------------------------------
def splitYearMonth( year_and_month )

	calendar = year_and_month.to_s
	return calendar.unpack("a4a2")	# 4文字 / 2文字に分割
end

#----------------------------------------------
# 西暦を平成に変換する
#----------------------------------------------
def convertYearNumberHeisei( year )

	# year を 1文字 / 3文字に分割 => 下三桁に12を加算
	year_number_array = year.unpack("a1a3")
	return ( year_number_array[1].to_i + 12 )
end

#----------------------------------------------
# 年月日から曜日を算出
#----------------------------------------------
def calcWeekDay( year, month, day )
	time = Time.mktime( year, month, day )
	return WDAYS[ time.wday ]
end

#----------------------------------------------
# 指定年月日が平日か（祝日考慮しない）
#----------------------------------------------
def isWeekday( year, month, day )

	w_day = calcWeekDay( year, month, day )
	if( w_day == "日" )      # 日曜
		return false
	elsif( w_day == "土" )   # 土曜
		return false
	else
		return true
	end
end

#----------------------------------------------
# 指定年月日が週末か（祝日考慮しない）
#----------------------------------------------
def isWeekend( year, month, day )

	if( isWeekday( year, month, day ) )
		return false
	else
		return true
	end
end

#----------------------------------------------
# @biref	指定範囲の年月を算出（外部からは呼ばない想定）
# @parm		start_time			開始生年月日（2000/01/01 想定）
# @parm		check_year_month	チェック年月（201501 想定）
#----------------------------------------------
def calcYearsImple( start_time, check_year_month )

	# 指定年月を算出
	str_calendar 	= splitYearMonth( check_year_month )
	year	 		= str_calendar[0].to_i
	month			= str_calendar[1].to_i
	time_now		= Date.new( year, month )

	# 生年月日を算出
	date_time_birth	= DateTime.parse( start_time )
	time_birth		= Date.new( date_time_birth.year, date_time_birth.mon, date_time_birth.day )

	# 年齢を算出
	diff			= time_now - time_birth
	result_age		= diff.to_f / 365

	return result_age;
end

#----------------------------------------------
# @biref	指定範囲の年月を算出
# @parm		start_time			開始生年月日（2000/01/01 想定）
# @parm		check_year_month	チェック年月（201501 想定）
#----------------------------------------------
def calcYears( start_time, check_year_month )
	range_year = calcYearsImple( start_time, check_year_month)
	return range_year.floor;
end

#----------------------------------------------
# @biref	指定範囲の年月文字列を算出（○歳○ヶ月ver）
# @parm		start_time			開始生年月日（2000/01/01 想定）
# @parm		check_year_month	チェック年月（201501 想定）
#----------------------------------------------
def calcAgeStrWithMonth( start_time, check_year_month )

	age_year		= calcAgeImple( start_time, check_year_month)
	result_month	= (age_year - age_year.floor) * 12
	result_age		= "#{result_year.floor}歳 #{result_month.ceil}ヶ月"

	return result_age;
end

#----------------------------------------------
# @biref	WIN32OLEの「cripting.FileSystemObject」のコピー機能利用
# @parm		src_path	コピー元のパス
# @parm		dst_path	コピー先のパス
#----------------------------------------------
def fsoCopyFile( src_path, dst_path )
	fso = WIN32OLE.new('Scripting.FileSystemObject')
	fso.CopyFile( src_path, dst_path )
end

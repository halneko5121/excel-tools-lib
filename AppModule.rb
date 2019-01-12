#
#　AppModule.rb
#
module AppModule

	#----------------------------------------------
	# アプリの main 関数
	# @parm	title	アプリタイトル
	# @parm	ver	アプリver
	# @parm	block	ブロック引数
	#----------------------------------------------
	def self.main( title, ver = nil, &block )
	
		title_param = nil
		if( ver == nil )
			title_param = "#{title}"
		else
			title_param = "#{title} Ver #{ver}"		
		end	
		system("title #{title_param}")

		puts "-----------------------------------------------------------"
		puts "start\n\n"
		
		block.call()

		puts "\nend"
		puts "-----------------------------------------------------------"
		successLogPrint()
	end
end

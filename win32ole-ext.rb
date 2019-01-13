require 'win32ole'

class WIN32OLE
    @const_defined = Hash.new

    # Excel オブジェクトの生成と定数の読み込み
    def WIN32OLE.new_with_const(prog_id, const_name_space)
        result = WIN32OLE.new(prog_id)

        # 二重定義防止
        unless @const_defined[const_name_space] then
            # WIN32OLE定数の読み込み
            # @note: WIN32OLE オブジェクトを生成しなければ定数を定義できない
            WIN32OLE.const_load(result, const_name_space)
            @const_defined[const_name_space] = true
        end
        return result
    end
end

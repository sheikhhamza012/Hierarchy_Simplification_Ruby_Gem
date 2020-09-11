
require 'rubyXL'
require 'rubyXL/convenience_methods'

class HierarchySimplification
  def self.generate_files(file_path)
    input = RubyXL::Parser.parse("file2.xlsx")
    nodes = []
    parents = []
    levels = []
    input[0].drop(1).each do |row| 
        i=row.size-1
        while i>0 do
            if row[i].value.nil?
                i-=1
                next
            end
            node=row[i].value
            level=input[0][0][i].value
            parent=""
            k=i-1
            while k>0 do 
                if !row[k].value.nil? 
                    if row[k].value ==row[i].value
                        level = input[0][0][k].value
                        k-=1
                        next
                    end
                    parent = row[k].value
                    break
                end
                k-=1
            end
            if nodes.each_index.select{|i| nodes[i] == node && parents[i] == parent}.size > 0 
                i-=1
                next
            end
            nodes.push(node)
            parents.push(parent)
            levels.push(level)
            i-=1;
        end
    end
    worksbook = RubyXL::Workbook.new
    worksheet= worksbook.first 
    worksheet.add_cell(0, 0, 'area') 
    worksheet.add_cell(0, 1, 'area_parent') 
    worksheet.add_cell(0, 2, 'Nivel') 
    worksheet.add_cell(0, 3, 'new_name') 
    nodes.each_with_index do |n,i|
        worksheet.add_cell(i+1, 0, n) 
        worksheet.add_cell(i+1, 1, parents[i]) 
        worksheet.add_cell(i+1, 2, levels[i]) 
    end
    worksbook.write("parents.xlsx")

    input[0].each do |row| 
      i=row.size-1
      while i>0 do
          if !row[i].value.nil?
              row[row.size-1].change_contents(row[i].value.to_s,row[row.size-1].formula)
              break
          end
          i-=1;
      end
    end
    input.write("file2.xlsx")

  end
end
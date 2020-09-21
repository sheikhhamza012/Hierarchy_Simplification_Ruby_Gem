
require 'rubyXL'
require 'rubyXL/convenience_methods'

class HierarchySimplification
  def self.generate_files(file_path)
    puts "reading file..."
    input = RubyXL::Parser.parse(file_path)
    nodes = []
    parents = []
    levels = []
    puts "finding parents"
    input[0].drop(1).each_with_index do |row,row_index| 
        puts "parent of row"+(row_index+1).to_s
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
    puts "writing to parents.xlsx"
    nodes.each_with_index do |n,i|
        worksheet.add_cell(i+1, 0, n) 
        worksheet.add_cell(i+1, 1, parents[i]) 
        worksheet.add_cell(i+1, 2, levels[i]) 
    end
    worksbook.write("parents.xlsx")
    
    puts "initiating second file..."
    input[0].drop(1).each_with_index do |row,row_index| 
        puts "analysing row: "+(row_index+1).to_s
        i=row.size-1
        while i>0 do
            if !row[i].value.nil?
                if row[input[0][0].size-1].nil?
                    input[0].add_cell(row_index+1,input[0][0].size-1,row[i].value)
                else
                    row[input[0][0].size-1].change_contents(row[i].value)
                end
                break
            end
            i-=1;
        end
    end
    puts "writing file..."
    input.write(file_path)

  end


  def self.rename_cells(file_path, parents_path)
        parents = []
        nodes = []
        new_names = []
        puts"opening file with new names.."
        input = RubyXL::Parser.parse(parents_path)
        
        puts"reading values to change.."
        input[0].drop(1).each_with_index do |row,row_index| 
            next if row[3]&.value.nil?
            nodes.push(row[0].value)
            parents.push(row[1].value)
            new_names.push(row[3].value)
        end
        
        puts"opening file to change.."
        input = RubyXL::Parser.parse(file_path)
        input[0].drop(1).each_with_index do |row,row_index| 
            i=input[0][0].size-1
            # byebug if row_index==4
            while i>0 do
                if !row[i]&.value.nil?
                    node = row[i].value
                    parent = row[i-1]&.value
                    k=i-1
                    while k>0
                            if !row[k]&.value.nil?
                                parent = row[k].value
                                break
                            end
                            k-=1
                    end
                    searched = nodes.each_index.select{|n| nodes[n] == node && parents[n] == parent}
                    if searched.size>0
                            puts "changing "+row[i].value.to_s+" to "+new_names[searched.first]
                            if row[i].value == row[input[0][0].size-1].value && i!=input[0][0].size-1
                                row[input[0][0].size-1].change_contents(new_names[searched.first])
                            end
                            value =row[i]&.value
                            k=i
                            while row[k]&.value == value
                                row[k].change_contents(new_names[searched.first])
                                k+=1
                            end
                        end
                end
                i-=1;
            end
        end
        puts"saving.."
        input.write(file_path)
        
    end
end
require_relative './domaci.rb'

sheet1 = ExcelParser.new("sheet.xlsx", 0)
sheet2 = ExcelParser.new("sheet.xlsx", 1)


 puts "Konvertovanje u matricu"
 p sheet1.to_matrix
 puts "------"

 puts "Pristupanje redu po indeksu (nulti red vraca red posle headera)"
 p sheet1.row(0)
 puts "------"

 puts "each metoda za tabelu"
 sheet1.each { |row| puts row }
 puts "------"

 puts "Direktan pristup koloni tabele"
 p sheet1['indeks']
 p sheet1['indeks'][2]
 puts "------"

 puts "Postavljanje vrednosti celiji u kolone"
 sheet1['indeks'][1] = "ri9322"
 p sheet1.print_all
 puts "------"

 puts "Pristup koloni po istoimenoj metodi"
 p sheet1.indeks.to_s
 puts "------"

 puts "Average i Sum metode kolone"
 p sheet1.field2.avg
 p sheet1.field2.sum
 puts "------"

# # Get row by column and value
 puts "Pristupanje redu po imenu kolone i vrednosti celije"
 p sheet1.indeks.rn323
 puts "------"

puts "Map, Select and Reduce metode nad kolonom"
  p sheet1.field2.map { |x| x * 2}
  p sheet1.field2.reduce(0) { |sum, x| sum + x }
  p sheet1.field2.select { |x| x > 3 }

require 'roo'

class ExcelParser
  include Enumerable

  def initialize(file_path, sheetnumber)
    @excel = Roo::Spreadsheet.open(file_path).sheet(sheetnumber)
    @data = parse_data
    create_dynamic_methods
  end

  def each(&block)
    # Određivanje maksimalnog broja redova
    max_length = @data.values.map(&:length).max
  
    # Iteriranje kroz svaki red
    max_length.times do |row_index|
      row_values = @data.keys.map { |key| @data[key][row_index] }
      block.call(row_values)
    end
  end

  def [](key)
    @data[key]
  end

  def []=(key, index, value)
    if @data[key] && index >= 0 && index < @data[key].size
      @data[key][index] = value
    end
  end

  # Metoda za štampanje
  def print_all
    max_length = @data.values.map(&:length).max
    max_length.times do |i|
      @data.keys.each { |key| print "#{@data[key][i] || ''}\t" }
      print "\n\n"
    end
  end

  def to_matrix
    headers = @data.keys
    matrix = [headers]
    max_length = @data.values.map(&:length).max
    max_length.times do |i|
      row = headers.map { |header| @data[header][i] }
      matrix << row
    end
    matrix
  end

  def row(index)
    return nil if index >= @data.values.first.length
    @data.keys.map { |key| @data[key][index] }
  end

  private

  def parse_data
    headers = @excel.row(1)
    data = Hash[headers.map {|header| [header, []]}]

    @excel.each_row_streaming(offset: 1) do |row|
      next if row_empty?(row)
      row.each_with_index do |cell, index|
        data[headers[index]] << cell.value
      end
    end

    data
  end

  def create_dynamic_methods
    @data.keys.each do |header|
      define_singleton_method(header.downcase.gsub(/\s+/, '_')) do
        DynamicColumn.new(self, header, @data[header])
      end
    end
  end

  def row_empty?(row)
    row.any? { |cell| cell_value_contains_total_or_subtotal?(cell) } ||
    row.all? { |cell| cell_nil_or_empty?(cell) }
  end
  
  def cell_value_contains_total_or_subtotal?(cell)
    cell_value = cell.value.to_s.strip.downcase
    cell_value.include?("total") || cell_value.include?("subtotal")
  end
  
  def cell_nil_or_empty?(cell)
    cell_value = cell.value.to_s.strip
    cell.nil? || cell_value.empty?
  end

  class DynamicColumn
     include Enumerable

    attr_reader :header, :values

    def initialize(parser, header, values)
      @parser = parser
      @header = header
      @values = values
    end

    def to_s
      @values
    end

    # Wrapper oko defaultne each metode
    def each(&block)
      @values.each(&block)
    end

    # Omogućavanje pristupa indeksu kao niz
    def [](index)
      @values[index]
    end

    # Dodeljivanje vrednosti
    def []=(index, value)
      @values[index] = value
    end

    def sum
      @values.compact.sum
    end
    
    def avg
      return 0 if @values.empty?
      @values.compact.sum.to_f / @values.compact.size
    end

    # Metode kao što su map, select i reduce će automatski raditi
    # Zahvaljujući uključivanju Enumerable modula i definiciji each metode

    # method_missing za podršku pozivanju metoda kao što su .rn9522
    def method_missing(method_name, *arguments, &block)
      value = method_name.to_s
      index = @values.index(value)
      return @parser.row(index) if index
      super
    end

  end

end

module Dossier
  class Xls

    HEADER = %Q{<?xml version="1.0" encoding="UTF-8"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">\n<Worksheet ss:Name="Sheet1">\n<Table>\n}
    FOOTER = %Q{</Table>\n</Worksheet>\n</Workbook>\n}

    def initialize(collection, headers = nil)
      @headers    = headers || collection.shift
      @collection = collection
    end

    def each 
      p = Axlsx::Package.new
      p.workbook.add_worksheet() do |sheet|
        sheet.add_row(@headers)
        @collection.each{|record| sheet.add_row(record)}
      end
      p.use_shared_strings = true
      p.serialize("circuit_id_validation.xlsx")
    end
  end
end

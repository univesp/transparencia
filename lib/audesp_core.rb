require 'xsd_populator'
require_relative 'portal_core'

def run_xml_verbas(worksheet)
  begin
    # OpenSSL::SSL::SSLContext::DEFAULT_PARAMS[:ciphers]=OpenSSL::SSL::SSLContext.new.ciphers;
    xsd_path = File.join(
                File.expand_path("../public/audesp/xml_dev/verbasremuneratorias",
                                 __FILE__)
                    .gsub('/lib',''),
                'AUDESP_VERBAS_REMUNERATORIAS_2020_A.XSD'
                )

    content = ''
    sheet = worksheet.sheets[5]
    sheet.rows.each_with_index do |row, index|
      next if index == 0 # header
      break if row[5].to_s.empty? # prevents empty lines

      content += "<VerbasRemuneratorias>
                    <Codigo>#{row[5].to_i.to_s}</Codigo>
                    <Nome>#{row[6].to_s}</Nome>
                    <EntraNoCalculoDoTetoConstitucional>
                      #{row[11].to_s.split('-')[0].strip}
                    </EntraNoCalculoDoTetoConstitucional>
                 </VerbasRemuneratorias>
                "
    end

    xsd_complete = XsdPopulator.new xsd: xsd_path
    File.open('output.txt','w') do |file|
      file.write xsd_complete
                   .populated_xml
                   .to_s
                   .gsub('<VerbasRemuneratorias>cvr:VerbasRemuneratorias_t</VerbasRemuneratorias>',content)
    end

    send_file "output.txt", :filename => 'UNIVESP.txt', :type =>
         'Application/octet-stream'
  rescue => e
    puts e.message
    puts e.backtrace
  end
end
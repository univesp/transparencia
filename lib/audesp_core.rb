require 'xsd_populator'
require_relative 'portal_core'

def run_xml_verbas(worksheet)
  begin
    template_path = File.join(
                      File.expand_path("../files/templates_audesp",
                                       __FILE__)
                          .gsub('/lib',''),
                      'verbas_remuneratorias.txt'
                      )

    template = File.read(template_path)

    items = ''
    maturities_sheet = worksheet.sheets[5]
    maturities_sheet.rows.each_with_index do |row, index|
      next if index == 0 # header
      break if row[5].to_s.empty? # prevents empty lines

      calculate_bool = row[11].to_s.split('-')[0].strip
      items += "<cap:VerbasRemuneratorias>
                    <cap:Codigo>#{row[5].to_i.to_s.strip}</cap:Codigo>
                    <cap:Nome>#{row[6].to_s.strip}</cap:Nome>
                    <cap:EntraNoCalculoDoTetoConstitucional
                      >#{calculate_bool}</cap:EntraNoCalculoDoTetoConstitucional>
                 </cap:VerbasRemuneratorias>
                "
    end

    header_sheet = worksheet.sheets[0]
    header_row = header_sheet.rows[1] # o cabeçalho é a linha zero

    template.gsub! '[VERBAS_REMUNERATORIAS]', items
    template.gsub! '[MUNICIPIO]', header_row[18].to_i.to_s.strip
    template.gsub! '[ENTIDADE]', header_row[19].to_i.to_s.strip
    template.gsub! '[ANO_EXERCICIO]', header_row[20].to_i.to_s.strip
    template.gsub! '[DATA_CRIACAO]', Time.now.strftime('%F')
    # TODO: verificar encode
    template.gsub! '[TIPO_DOCUMENTO]',
                   'Cadastro de Verbas Remuneratórias'

    File.open('output.xml','w') do |file|
      file.write template
    end

    send_file "output.xml",
              filename: 'UNIVESP.xml',
              type: 'Application/octet-stream'
  rescue => e
    puts e.message
    puts e.backtrace
  end
end
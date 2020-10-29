MUNICIPIO = 7107.freeze
ENTIDADE = 11219.freeze
ROOT_PATH = File.expand_path('../files/templates_audesp',
                             __FILE__)
                .gsub('/lib','')
                .freeze

def run_xml_folha_ordinaria(worksheet)
  begin
    template_master = File.read(
                        File.join(ROOT_PATH, 'folha_ordinaria.txt')
                      )
    template_agent = File.read(
        File.join(ROOT_PATH, 'identificacao_agente_publico.txt')
    )
    template_remuneration = File.read(
        File.join(ROOT_PATH, 'verbas.txt')
    )

    financial_list = []
    financial_data_sheet = worksheet.sheets[3]
    financial_data_sheet.rows.each_with_index do |row, index|
      person_code = row[9].to_i


      launches = [] # controla cada lançamento financeiro

      index = 11
      (0..30).each do
        ## Código da verba
        code = row[index].to_s.to_i
        break if code == 0 # se o código não existe não há mais nenhum
        # lançamento daqui pra frente

        ## Valor
        value = ((row[index + 1].to_f).fdiv(100)).to_s

        ## Espécie
        # AUDESP: 1 - Desconto,2 - Vencimento
        # Portal: D - Débito, C - Crédito ao funcionário
        #         R - Redutor Salarial / Reposição (não utilizado),
        #         S - Reposição de Desconto (não utilizado)
        kind = row[index + 2].to_s.split('-')[0].to_s.strip
        kind = kind == 'C' ? 2 : 1

        ## Natureza
        # AUDESP: 1 - Atrasado, 2 - Normal
        # Portal: M - Do Mês, A - Atrasado
        nature = row[index + 3].to_s.split('-')[0].to_s.strip
        nature = nature == 'A' ? 1 : 2

        ## Código do tipo da verba
        type_code = row[index + 5].to_s.split('-')[0].to_s.strip.to_i

        index += 8 # incremento de 8 para pular para o próximo bloco de
        # lançamentos

        launches.push({
          code: code,
          value: value,
          nature: nature,
          kind: kind,
          type_code: type_code
        })
      end

      financial_list.push({
                              person_code: person_code,
                              launches: launches
                          })
    end


    items = ''
    personal_data_sheet = worksheet.sheets[2]
    personal_data_sheet.rows.each_with_index do |row, index|
      next if index == 0 # pula o cabeçalho
      break if row[2].to_s.empty? # evita as linhas em branco

      person_financial_data = financial_list.select { |f|
        f[:person_code] == row[9].to_i
      }

      remuneration = ''
      person_financial_data.first[:launches].each do |launch|
        remuneration += template_remuneration
                          .gsub('[MUNICIPIO]',MUNICIPIO.to_s)
                          .gsub('[ENTIDADE]',ENTIDADE.to_s)
                          .gsub('[CODIGO_VERBA]', launch[:code].to_s)
                          .gsub('[VALOR]', launch[:value].to_s)
                          .gsub('[NATUREZA]', launch[:nature].to_s)
                          .gsub('[ESPECIE]', launch[:kind].to_s)
                          .gsub('[CODIGO_TIPO_VERBA]', launch[:type_code].to_s)
      end if person_financial_data.size > 0

      items += template_agent
                 .gsub('[CPF]', mount_cpf(row))
                 .gsub('[NOME]',"#{row[13].to_s.strip}")
                 .gsub('[MUNICIPIO]',"#{row[146].to_s.strip}")
                 .gsub('[ENTIDADE]',"#{row[147].to_s.strip}")
                 .gsub('[CARGO_POLITICO]',"#{row[148].to_s
                                               .split('-')[0]
                                               .to_s
                                               .strip}")
                 .gsub('[FUNCAO_GOVERNO]',"#{row[149].to_s.strip}")
                 .gsub('[ESTAGIARIO]',"#{row[150].to_s
                                             .split('-')[0]
                                             .to_s
                                             .strip}")
                 .gsub('[CODIGO_CARGO]',"#{row[45].to_s.strip}")
                 .gsub('[SITUACAO]',"#{row[151].to_s
                                                .split('-')[0]
                                                .to_s
                                                .strip}")
                 .gsub('[REGIME_JURIDICO]',"#{row[152].to_s
                                                  .split('-')[0]
                                                  .to_s
                                                  .strip}")
                 .gsub('[AUTORIZACAO_TETO]',"#{row[153].to_s
                                                   .split('-')[0]
                                                   .to_s
                                                   .strip}")
                 .gsub('[REMUNERACAO_BRUTA]',"#{float_from_text(row[131])}")
                 .gsub('[DESCONTOS]',"#{float_from_text(row[132])}")
                 .gsub('[REMUNERACAO_LIQUIDA]',"#{float_from_text(row[133])}")
                 .gsub('[VERBAS]',remuneration)


    end

    header_sheet = worksheet.sheets[0]
    header_row = header_sheet.rows[1] # o cabeçalho é a linha zero

    template_master.gsub! '[VERBAS_REMUNERATORIAS]', items
    template_master.gsub! '[MUNICIPIO]', header_row[18].to_i.to_s.strip
    template_master.gsub! '[ENTIDADE]', header_row[19].to_i.to_s.strip
    template_master.gsub! '[ANO_EXERCICIO]', header_row[20].to_i.to_s.strip
    template_master.gsub! '[MES_EXERCICIO]', header_row[21].to_i.to_s.strip
    template_master.gsub! '[DATA_CRIACAO]', Time.now.strftime('%F')
    template_master.gsub! '[IDENTIFICACAO_AGENTE]', items

    save_and_send_file template_master,
                       :folha_ordinaria

  rescue => e
    puts e.message
    puts e.backtrace
  end
end

def run_xml_pagamento_folha_ordinaria(worksheet)
  begin
    template_folha_ordinaria = File.read(
      File.join(ROOT_PATH, 'pagamento_folha_ordinaria.txt')
    )
    fragment_identificacao_agente = File.read(
      File.join(ROOT_PATH, 'fragmentos/identificacao_agente_folha.txt')
    )

    agents = ''
    personal_data_sheet = worksheet.sheets[2]
    personal_data_sheet.rows.each_with_index do |row, index|
      next if index == 0 # pula o cabeçalho
      break if row[2].to_s.empty? # evita as linhas em branco

      ## Forma de pagamento
      # 1 - Conta corrente;2 - Demais;3 - Ambas

      agents += fragment_identificacao_agente
                  .gsub('[CPF]', mount_cpf(row))
                  .gsub('[FORMA_PAGAMENTO]',"#{first_from_split(row[156])}")
                  .gsub('[NUMERO_BANCO]',"#{row[157].to_s.strip}")
                  .gsub('[AGENCIA]',"#{row[158].to_s.strip}")
                  .gsub('[CONTA_CORRENTE]',"#{row[159].to_s.strip}")
                  .gsub('[PAGO_CC]',"#{float_from_text(row[160])}")
                  .gsub('[PAGO_OUTROS]',"#{float_from_text(row[161])}")
                  .gsub('[MUNICIPIO]', MUNICIPIO.to_s)
                  .gsub('[ENTIDADE]', ENTIDADE.to_s)
    end

    header_sheet = worksheet.sheets[0]
    header_row = header_sheet.rows[1] # o cabeçalho é a linha zero

    template_folha_ordinaria.gsub! '[MUNICIPIO]', MUNICIPIO.to_s
    template_folha_ordinaria.gsub! '[ENTIDADE]', ENTIDADE.to_s
    template_folha_ordinaria.gsub! '[ANO_EXERCICIO]',
                                   header_row[20].to_i.to_s.strip
    template_folha_ordinaria.gsub! '[MES_EXERCICIO]',
                                   header_row[21].to_i.to_s.strip
    template_folha_ordinaria.gsub! '[DATA_CRIACAO]',
                                   Time.now.strftime('%F')
    # TODO: verificar se ano/mês de pagamento é o mesmo do exercício
    template_folha_ordinaria.gsub! '[ANO_PAGAMENTO]',
                                   header_row[20].to_i.to_s.strip
    template_folha_ordinaria.gsub! '[MES_PAGAMENTO]',
                                   header_row[21].to_i.to_s.strip
    template_folha_ordinaria.gsub! '[IDENTIFICACAO_AGENTES]', agents


    save_and_send_file template_folha_ordinaria,
                  :pagamento_folha_ordinaria
  rescue => e
  end
end

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

    save_and_send_file template,
                :verbas_remuneratorias
  rescue => e
    File.open('exceptions.log','a') do |file|
      file << "#{Time.now.strftime('%FT%T%:z')}
                \n\t#{e.message}
                \n\t#{e.backtrace}"
    end
  end
end

# --- AUXILIARES --- #

###
# Obtem a primeira parte do texto com separador de uma célula.
def first_from_split(cell)
  cell.to_s
    .split('-')[0]
    .to_s
    .strip
end

###
# Converte os valores vindos da planilha do Portal sem separador de decimal e
# os formata de acordo com o padrão da AUDESP.
def float_from_text(cell)
  '%.2f' % (cell.to_f).fdiv(100)
end

###
# Monta o CPF o qual consta na planilha do Portal em duas células: número
# (coluna 24) e controle (coluna 25).
def mount_cpf(row)
  "#{row[24].to_s.strip}#{row[25].to_s.strip}"
end

###
# Salva um arquivo físico e o oferece para download no browser
def save_and_send_file(file_content, origin)
  file_name = case origin
              when :folha_ordinaria
                'UNIVESP_Folha_Ordinaria.xml'
              when :pagamento_folha_ordinaria
                'UNIVESP_Pagamento_Folha_Ordinaria.xml'
              when :verbas_remuneratorias
                'UNIVESP_Verbas_Remuneratorias.xml'
              end

  File.open(file_name, 'w') do |file|
    file.write file_content
  end

  send_file "output.xml",
            filename: file_name,
            type: 'Application/octet-stream'
end
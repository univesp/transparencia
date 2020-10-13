require 'sinatra'
require 'simple_xlsx_reader'
require 'cgi'
require_relative 'lib/core'

get '/' do
  erb :index
end

post '/upload' do
  file = params[:file]

  return CGI.unescape("<p>
            Nenhum arquivo foi carregado
          </p><br>
          <a href=\"/portal-da-transparencia\">Voltar</a>") unless file

  allowed_types = %w(application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
  if allowed_types.include? file[:type]
    # Leitura da planilha carregada
    worksheet = SimpleXlsxReader.open file[:tempfile]

    # Lendo DADOS DE ENCARGOS SOCIAIS
    hash_charge_data = charge_data worksheet.sheets[1]

    # Lendo DADOS PESSOAIS
    hash_personal_data = personal_data worksheet.sheets[2]

    # Lendo DADOS FINANCEIROS
    hash_financial_data = financial_data worksheet.sheets[3]

    # Lendo TABELA DE CARGOS E FUNÇÕES
    hash_job_data = job_data worksheet.sheets[4]

    # Lendo TABELA DE VENCIMENTOS E DESCONTOS
    hash_maturities_data = maturities_data worksheet.sheets[5]

    # Lendo TABELA DE FAIXAS SALARIAIS
    hash_salary_range_data = salary_range_data worksheet.sheets[6]

    # Lendo TABELA DE UNIDADES ADMINISTRATIVAS
    hash_unit_data = unit_data worksheet.sheets[7]

    # Montando o HEADER
    hash_header = header_data(worksheet.sheets[0],
                              hash_charge_data, hash_personal_data,
                              hash_financial_data, hash_job_data,
                              hash_maturities_data, hash_salary_range_data,
                              hash_unit_data)

    write_file('output.txt' ,
               hash_header, hash_charge_data, hash_personal_data,
               hash_financial_data, hash_job_data, hash_maturities_data,
               hash_salary_range_data, hash_unit_data)

    send_file "output.txt", :filename => 'UNIVESP.txt', :type => 'Application/octet-stream'
  else
    return CGI.unescape("<p>
            Apenas arquivos com extensão XLSX são permitidos
          </p><br>
          <a href=\"/portal-da-transparencia\">Voltar</a>") unless file
  end
end
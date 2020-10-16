require 'sinatra'
require 'simple_xlsx_reader'
require 'cgi'
require_relative 'lib/audesp_core'
require_relative 'lib/portal_core'

get '/' do
  erb :index
end

post '/upload' do
  file = params[:file]
  conversion_type = params[:conversion_type]

  return CGI.unescape("<p>
            Nenhum arquivo foi carregado
          </p><br>
          <a href=\"#{url('/')}\">Voltar</a>") unless file

  allowed_types = %w(application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
  if allowed_types.include? file[:type]
    worksheet = SimpleXlsxReader.open file[:tempfile]
    case conversion_type
    when 'txt_portal'
      run_txt_portal worksheet
    when 'xml_verbas'
      run_xml_verbas worksheet
    when 'xml_folha_ordinaria'
      run_xml_folha_ordinaria worksheet
    when 'xml_pagamento_folha_ordinaria'
      run_xml_pagamento_folha_ordinaria worksheet
    end
  else
    return CGI.unescape("<p>
            Apenas arquivos com extensão XLSX são permitidos
          </p><br>
          <a href=\"/portal-da-transparencia\">Voltar</a>") unless file
  end
end
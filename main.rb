require 'sinatra'
require 'sinatra/basic_auth'
require 'simple_xlsx_reader'
require 'cgi'
require 'yaml'
require_relative 'lib/audesp_core'
require_relative 'lib/portal_core'

authorize do |username, password|
  users = YAML.load(File.read('files/users/users.yml'))
  current_user = users.select { |u|
    u['username'].to_s == username.to_s &&
    u['password'].to_s == password.to_s
  }
  current_user.size == 1
end

def upload
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
    when 'xml_resumo_mensal'
      run_xml_resumo_mensal worksheet
    end
  else
    return CGI.unescape("<p>
            Apenas arquivos com extensão XLSX são permitidos
          </p><br>
          <a href=\"/portal-da-transparencia\">Voltar</a>") unless file
  end
end

protect do
  get '/' do
    erb :index
  end

  post '/upload' do
    return upload
  end

  post '/transparencia/upload' do
    return upload
  end
end
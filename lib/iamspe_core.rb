#!/bin/env ruby
# encoding: utf-8
# This script contains the business rules used to generate the file
# to be sent to PRODESP.

require_relative 'utils'


CLIENTE='306'.freeze
IDENT_ENTID='F'.freeze
IDENT_UA='95694'.freeze
NOME_ENTIDADE='FUND.UN.VIRTUAL EST.SP UNIVESP'.freeze
IDENT_ORGAO='10'.freeze
IDENT_UO='46'.freeze
IDENT_UD='30'.freeze


def run_txt_iamspe(worksheet)
  File.open('output.txt', 'w') do |f|
    f.write personal_data_iamspe worksheet.sheets[2]
  end

  send_file "output.txt",
            filename: 'UNIVESP_IAMSPE.txt',
            type: 'Application/octet-stream'
end

def personal_data_iamspe(sheet)
  content = ''
  sheet.rows.each_with_index do |row, index|
    next if index == 0 # cabeçalho

    # IAMSPE - ADERIU?
    next if row[163].to_s != '1 - SIM' # pula a linha se a opção
    # IAMSPE - ADERIU?  diferente de "1 - SIM"

    line = ''

    # CLIENTE n3
    line += CLIENTE.pad_with_right_spaces! 3 # [FJ]

    # IDENT_FUNCIO n11
    line += "#{row[9]}".pad_with_right_spaces! 11 # [J]

    # SUB_IDENT_FUNCIO n2
    line += "#{row[10]}".pad_with_right_spaces! 2 # [K]

    # NOME_FUNCIONARIO a30
    line += "#{row[13]}".pad_with_right_spaces! 30 # [N]

    # NUMERO_IDENTIDADE a11
    line += ''.pad_with_right_spaces! 11

    # NUMERO_CPF n9
    line += "#{row[24]}".pad_with_right_spaces! 9 # [Y]

    # CONTROLE_CPF n2
    line += "#{row[25]}".pad_with_right_spaces! 2 # [Z]

    # SEXO a1
    line += "#{row[15]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [P]

    # EST_CIVIL a1
    line += "#{row[16]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [Q]

    # DATA_NASC n8
    line += "#{row[14]}".pad_with_right_spaces! 8 # [O]

    # POP_CAR a2
    line += "#{row[164]}".pad_with_right_spaces! 2 # [FJ]

    # COD_SITUA_FUNCIO a2
    line += "#{row[165]}".split(' = ')[0].to_s.pad_with_right_spaces! 2 # [FK]

    # MOTIVO_SITUA_FUNCIO n3
    line += "#{row[166]}".split(' = ')[0].to_s.pad_with_right_spaces! 3 # [FL]

    # DAT_INIC_SITUA_FUNCIO n8
    line += "#{row[61]}".pad_with_right_spaces! 8 # [BJ]

    # VD_IAMSPE n15
    line += "#{comma_for_currencies row[167]}".pad_with_right_spaces! 15 # [FM]

    # VD_AGREGADO n15
    line += "#{comma_for_currencies row[168]}".pad_with_right_spaces! 15 # [FM]

    # IDENT_ENTID a1
    line += IDENT_ENTID.pad_with_right_spaces! 1 # [FN]

    # IDENT_UA n7
    line += IDENT_UA.pad_with_right_spaces! 7

    # NOME_ENTIDADE a40
    line += NOME_ENTIDADE.pad_with_right_spaces! 40

    # IDENT_ORGAO n2
    line += IDENT_ORGAO.pad_with_right_spaces! 2

    # IDENT_UO n3
    line += IDENT_UO.pad_with_right_spaces! 3

    # VD_FÉRIAS n15
    line += "#{comma_for_currencies row[169]}".pad_with_right_spaces! 15 # [FO]

    # VD_PARCELAMENTO n15
    line += "#{comma_for_currencies row[170]}".pad_with_right_spaces! 15 # [FP]

    # VD_RESERVA n15
    line += "#{comma_for_currencies row[171]}".pad_with_right_spaces! 15 # [FQ]

    # VD_IAMSPE_LEI_11456 n15
    line += "#{comma_for_currencies row[172]}".pad_with_right_spaces! 15 # [FR]

    # VD_AGREGADOS_LEI_11456 n15
    line += "#{comma_for_currencies row[173]}".pad_with_right_spaces! 15 # [FS]

    # VD_FÉRIAS_LEI_11456 n15
    line += "#{comma_for_currencies row[174]}".pad_with_right_spaces! 15 # [FT]

    # VD_AFAST_IAMSPE_CEETEPS n15
    line += "#{comma_for_currencies row[175]}".pad_with_right_spaces! 15 # [FU]

    # VD_AGREG_IAMSPE_CEETEPS n15
    line += "#{comma_for_currencies row[176]}".pad_with_right_spaces! 15 # [FV]

    # IDENT_MUNICIPIO n5
    line += ''.pad_with_right_spaces! 5

    # DENOM_NACIO a20
    line += ''.pad_with_right_spaces! 20

    # NOME_MAE a30
    line += "#{row[19]}".pad_with_right_spaces! 30 # [T] C 30

    # EST_EMIS_IDENTIDADE a2
    line += ''.pad_with_right_spaces! 2

    # NUMERO_PISPASEP n11
    line += ''.pad_with_right_spaces! 11

    # DATA_INGR_SERV_PUBL n8
    line += ''.pad_with_right_spaces! 8

    # ANO_PRIM_EMPREGO n4
    line += ''.pad_with_right_spaces! 4

    # COD_BANCO n5
    line += ''.pad_with_right_spaces! 5

    # COG_AGENCIA n5
    line += ''.pad_with_right_spaces! 5

    # DIG_AGENCIA n1
    line += ''.pad_with_right_spaces! 1

    # NUMERO_CONTA_CORRENTE n7
    #conta_cc = row[159].to_s[0..-2].to_s
    #conta_dig = row[159].to_s[-1].to_s
    #line += conta_cc.pad_with_right_spaces! 7
    line += ''.pad_with_right_spaces! 7

    # DIGITO_CONTA_CORRENTE n1
    #line += conta_dig.pad_with_right_spaces! 1
    line += ''.pad_with_right_spaces! 1

    # ENDERECO_RESIDE a50
    line += "#{row[35]}".pad_with_right_spaces! 50 # [AJ]

    # ENDERECO_BAIRRO a30
    line += ''.pad_with_right_spaces! 30

    # CEP_RESIDENCIAL n8
    line += "#{row[36]}".pad_with_right_spaces! 8 # [AK]

    # CIDADE_RESIDENCIAL a30
    line += "#{row[37]}".pad_with_right_spaces! 30 # [AL]

    # VD_IAMSPE_SPPREV_070127 n15
    line += "#{comma_for_currencies row[177]}".pad_with_right_spaces! 15 # [FW]

    # VD_IAMSPE_SPPREV_700062 n15
    line += "#{comma_for_currencies row[178]}".pad_with_right_spaces! 15 # [FX]

    # VD_IAMSPE_SPPREV_700371 n15
    line += "#{comma_for_currencies row[179]}".pad_with_right_spaces! 15 # [FY]

    # VD_IAMSPE_SPPREV_700372 n15
    line += "#{comma_for_currencies row[180]}".pad_with_right_spaces! 15 # [FZ]

    # VD_IAMSPE_SPPREV_000504 n15
    line += "#{comma_for_currencies row[181]}".pad_with_right_spaces! 15 # [GA]

    # VD_IAMSPE_ODONTO_SEFAZ n15
    line += "#{comma_for_currencies row[182]}".pad_with_right_spaces! 15 # [GB]

    # VD_IAMSPE_ODONTO_SPPREV n15
    line += "#{comma_for_currencies row[183]}".pad_with_right_spaces! 15 # [GC]

    # VD_IAMSPE_BENEFI_070119 n15
    line += "#{comma_for_currencies row[184]}".pad_with_right_spaces! 15 # [GD]

    # VD_IAMSPE_AGRESF_070120 n15
    line += "#{comma_for_currencies row[185]}".pad_with_right_spaces! 15 # [GE]

    # VD_IAMSPE_BENESF_070121 n15
    line += "#{comma_for_currencies row[186]}".pad_with_right_spaces! 15 # [GF]

    # VD_IAMSPE_13SAL_070122 n15
    line += "#{comma_for_currencies row[187]}".pad_with_right_spaces! 15 # [GG]

    # VD_IAMSPE_AGRE13_070123 n15
    line += "#{comma_for_currencies row[188]}".pad_with_right_spaces! 15 # [GH]

    # VD_IAMSPE_BENE13_070124 n15
    line += "#{comma_for_currencies row[189]}".pad_with_right_spaces! 15 # [GI]

    # VD_IAMSPE_DEPENT_070641 n15
    line += "#{comma_for_currencies row[190]}".pad_with_right_spaces! 15 # [GJ]

    # VD_IAMSPE_070642 n15
    line += "#{comma_for_currencies row[191]}".pad_with_right_spaces! 15 # [GK]

    # VD_IAMSPE_070037 n15
    line += "#{comma_for_currencies row[192]}".pad_with_right_spaces! 15 # [GL]

    # VD_IAMSPE_907100 n15
    line += "#{comma_for_currencies row[193]}".pad_with_right_spaces! 15 # [GM]

    # VD_IAMSPE_901699 n15
    line += "#{comma_for_currencies row[194]}".pad_with_right_spaces! 15 # [GN]

    # VD_IAMSPE_000924 n15
    line += "#{comma_for_currencies row[195]}".pad_with_right_spaces! 15 # [GO]

    # VD_IAMSPE_070006 n15
    line += "#{comma_for_currencies row[196]}".pad_with_right_spaces! 15 # [GP]

    # VD_IAMSPE_000510 n15
    line += "#{comma_for_currencies row[197]}".pad_with_right_spaces! 15 # [GQ]

    # VD_IAMSPE_907015 n15
    line += "#{comma_for_currencies row[198]}".pad_with_right_spaces! 15 # [GR]

    # VD_IAMSPE_930401 n15
    line += "#{comma_for_currencies row[199]}".pad_with_right_spaces! 15 # [GS]

    # VD_IAMSPE_000110 n15
    line += "#{comma_for_currencies row[200]}".pad_with_right_spaces! 15 # [GT]

    # VD_IAMSPE_900302 n15
    line += "#{comma_for_currencies row[201]}".pad_with_right_spaces! 15 # [GU]

    # VD_IAMSPE_000923 n15
    line += "#{comma_for_currencies row[202]}".pad_with_right_spaces! 15 # [GV]

    # VD_IAMSPE_000826 n15
    line += "#{comma_for_currencies row[203]}".pad_with_right_spaces! 15 # [GW]

    # IDENT_UD n3
    line += IDENT_UD.pad_with_right_spaces! 3 # [GW]

    # VD_IAMSPE_070125 n15
    line += "#{comma_for_currencies row[204]}".pad_with_right_spaces! 15 # [GW]

    # VD_IAMSPE_070126 n15
    line += "#{comma_for_currencies row[205]}".pad_with_right_spaces! 15 # [GW]

    line += "\n"
    content += line
  end
  content
end

def comma_for_currencies(value)
  return '' if value.to_s.empty?

  # garante que o valor tenha sempre 3 dígitos (previne a vírgula quando valor
  # for menos de R$ 1,00)
  hundred_value =  value.to_s.rjust(3,'0')
  hundred_value.insert hundred_value.size-2, ','
end
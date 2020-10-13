#!/bin/env ruby
# encoding: utf-8
# This script contains the business rules used to generate the file
# to be sent to PRODESP.

require_relative 'utils'

# Reads the data in the "Detalhe - ENCARGOS SOCIAIS  BEN" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def charge_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    if seq == 0 # jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[BG] Z 15 means that data in column BG must be of only-numeric type and
    #        having 15 characters
    line = ""
    line += "ES".pad_with_right_spaces! 2 # [A] C 02
    line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "".pad_with_right_spaces! 33 # [E] C 33
    line += "#{row[5]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [F] Z 06
    line += "#{row[6]}".pad_with_left_zeroes! 15 # [G] Z 15
    line += "#{row[7]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [H] Z 06
    line += "#{row[8]}".pad_with_left_zeroes! 15 # [I] Z 15
    line += "#{row[9]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [J] Z 06
    line += "#{row[10]}".pad_with_left_zeroes! 15 # [K] Z 15
    line += "#{row[11]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [L] Z 06
    line += "#{row[12]}".pad_with_left_zeroes! 15 # [M] Z 15
    line += "#{row[13]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [N] Z 06
    line += "#{row[14]}".pad_with_left_zeroes! 15 # [O] Z 15
    line += "#{row[15]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [P] Z 06
    line += "#{row[16]}".pad_with_left_zeroes! 15 # [Q] Z 15
    line += "#{row[17]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [R] Z 06
    line += "#{row[18]}".pad_with_left_zeroes! 15 # [S] Z 15
    line += "#{row[19]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [T] Z 06
    line += "#{row[20]}".pad_with_left_zeroes! 15 # [U] Z 15
    line += "#{row[21]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [V] Z 06
    line += "#{row[22]}".pad_with_left_zeroes! 15 # [W] Z 15
    line += "#{row[23]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [X] Z 06
    line += "#{row[24]}".pad_with_left_zeroes! 15 # [Y] Z 15
    line += "#{row[25]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [Z] Z 06
    line += "#{row[26]}".pad_with_left_zeroes! 15 # [AA] Z 15
    line += "#{row[27]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AB] Z 06
    line += "#{row[28]}".pad_with_left_zeroes! 15 # [AC] Z 15
    line += "#{row[29]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AD] Z 06
    line += "#{row[30]}".pad_with_left_zeroes! 15 # [AE] Z 15
    line += "#{row[31]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AF] Z 06
    line += "#{row[32]}".pad_with_left_zeroes! 15 # [AG] Z 15
    line += "#{row[33]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AH] Z 06
    line += "#{row[34]}".pad_with_left_zeroes! 15 # [AI] Z 15
    line += "#{row[35]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AJ] Z 06
    line += "#{row[36]}".pad_with_left_zeroes! 15 # [AK] Z 15
    line += "#{row[37]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AL] Z 06
    line += "#{row[38]}".pad_with_left_zeroes! 15 # [AM] Z 15
    line += "#{row[39]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AN] Z 06
    line += "#{row[40]}".pad_with_left_zeroes! 15 # [AO] Z 15
    line += "#{row[41]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AP] Z 06
    line += "#{row[42]}".pad_with_left_zeroes! 15 # [AQ] Z 15
    line += "#{row[43]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AR] Z 06
    line += "#{row[44]}".pad_with_left_zeroes! 15 # [AS] Z 15
    line += "#{row[45]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AT] Z 06
    line += "#{row[46]}".pad_with_left_zeroes! 15 # [AU] Z 15
    line += "#{row[47]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AV] Z 06
    line += "#{row[48]}".pad_with_left_zeroes! 15 # [AW] Z 15
    line += "#{row[49]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AX] Z 06
    line += "#{row[50]}".pad_with_left_zeroes! 15 # [AY] Z 15
    line += "#{row[51]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [AZ] Z 06
    line += "#{row[52]}".pad_with_left_zeroes! 15 # [BA] Z 15
    line += "#{row[53]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BB] Z 06
    line += "#{row[54]}".pad_with_left_zeroes! 15 # [BC] Z 15
    line += "#{row[55]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BD] Z 06
    line += "#{row[56]}".pad_with_left_zeroes! 15 # [BE] Z 15
    line += "#{row[57]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BF] Z 06
    line += "#{row[58]}".pad_with_left_zeroes! 15 # [BG] Z 15
    line += "#{row[59]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BH] Z 06
    line += "#{row[60]}".pad_with_left_zeroes! 15 # [BI] Z 15
    line += "#{row[61]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BJ] Z 06
    line += "#{row[62]}".pad_with_left_zeroes! 15 # [BK] Z 15
    line += "#{row[63]}".split(' ')[0].to_s.pad_with_left_zeroes! 6 # [BL] Z 06
    line += "#{row[64]}".pad_with_left_zeroes! 15 # [BM] Z 15
    line += "".pad_with_right_spaces! 72 # [BN] C 73
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content, :qtd => seq }
end

# Reads the data in the "Detalhe - Dados FINANCEIROS" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def financial_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    begin

      if seq == 0 # jumps the header
        seq += 1
        next
      end

      break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

      # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
      #                                      numeric)	LIMIT SIZE
      # Example:
      # 	[BG] C 01 means that data in column BG must be of alphanumeric type and
      #        having 1 character
      line = ""
      line += "F2".pad_with_right_spaces! 2 # [A] C 02
      line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
      line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
      line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
      line += "#{row[4]}".pad_with_left_zeroes! 3 # [E] Z 03
      line += "#{row[5]}".pad_with_left_zeroes! 7 # [F] Z 07
      line += "".pad_with_right_spaces! 5 # [G] C 05
      line += "#{row[7]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [H] Z 02
      line += "#{row[8]}".split(' ')[0].to_s.pad_with_left_zeroes! 4 # [I] Z 04
      line += "#{row[9]}".pad_with_left_zeroes! 10 # [J] Z 10
      line += "#{row[10]}".pad_with_left_zeroes! 2 # [K] Z 02

      # a estrutura a seguir se repete por 30 entradas
      index = 11
      (0..30).each do
        line += "#{row[index]}".pad_with_left_zeroes! 6 # [L] Z 06
        line += "#{row[index + 1]}".pad_with_left_zeroes! 13 # [M] Z 13
        line += "#{row[index + 2]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [N] C 01
        line += "#{row[index + 3]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [O] C 01
        line += "#{row[index + 4]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [P] C 01
        # row[index + 5] Campo AUDESP
        # row[index + 6] Campo AUDESP
        # row[index + 7] Campo AUDESP
        index += 8
      end

      line += "".pad_with_right_spaces! 18 # [JA] C 19
      line += "\n" # considered as one char

      puts line

      content += line
      seq += 1
    rescue => e
      puts e.inspect
    end


  end

  { :content => content, :qtd => seq }
end

# Reads the data in the "Header" tab of worksheet and then converts them to
# lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @param [Hash] hash_charge_data Hash that contains all the data of
#                                "Detalhe - ENCARGOS SOCIAIS  BEN" tab
# @param [Hash] hash_personal_data Hash that contains all the data of
#                                  "Detalhe - Dados PESSOAIS  FUNCI" tab
# @param [Hash] hash_financial_data Hash that contains all the data of
#                                   "Detalhe - Dados FINANCEIROS" tab
# @param [Hash] hash_job_data Hash that contains all the data of
#                             "Detalhe - Tabela de CARGO  FUNÇ" tab
# @param [Hash] hash_maturities_data Hash that contains all the data of
#                                    "Detalhe - Tabela de VENCIMENTO" tab
# @param [Hash] hash_salary_range_data Hash that contains all the data of
#                                      "Detalhe - Tabela de FAIXA SALAR" tab
# @param [Hash] hash_unit_data Hash that contains all the data of
#                              "Detalhe - Tabela de UNIDADES AD" tab
# @return [Hash] Content of the tab converted to text line
def header_data(sheet, hash_charge_data, hash_personal_data,
                hash_financial_data, hash_job_data, hash_maturities_data,
                hash_salary_range_data, hash_unit_data)
  content = '' # content of text line
  seq = 0 # sequence counter

  # sums the amount of registers of all the worksheet tabs
  counter_arr = [
      hash_charge_data[:qtd] - 1,
      hash_personal_data[:qtd] - 1,
      hash_financial_data[:qtd] - 1,
      hash_job_data[:qtd] - 1,
      hash_maturities_data[:qtd] - 1,
      hash_salary_range_data[:qtd] - 1,
      hash_unit_data[:qtd] - 1
  ]
  counter = counter_arr.inject(:+)

  sheet.rows.each do |row|
    if seq == 0  # jumps the header
      seq += 1
      next
    end

    break if "#{row[1]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[Q] C 553 means that data in column Q must be of only-numeric type and
    #       having 553 characters
    line = ""
    line += "".pad_with_right_spaces! 9 # [A] C 09
    line += "#{row[1]}".pad_with_left_zeroes! 2 # [B] Z 02
    line += "#{row[2]}".pad_with_left_zeroes! 3 # [C] Z 03
    line += "".pad_with_right_spaces! 33 # [D] C 33
    line += "#{row[4]}".pad_with_left_zeroes! 6 # [E] Z 06
    line += "#{counter}".pad_with_left_zeroes! 11 # [F] Z 11
    line += "#{(hash_personal_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [G] Z 11
    line += "#{(hash_financial_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [H] Z 11
    line += "#{(hash_job_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [I] Z 11
    line += "".pad_with_right_spaces! 11 # [J] C 11
    line += "#{(hash_maturities_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [K] Z 11
    line += "#{(hash_salary_range_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [L] Z 11
    line += "#{hash_personal_data[:total_gross]}".pad_with_left_zeroes! 15 # [M] Z 15
    line += "#{hash_personal_data[:total_discounts]}".pad_with_left_zeroes! 15 # [N] Z 15
    line += "#{hash_personal_data[:total_net]}".pad_with_left_zeroes! 15 # [N] Z 15
    line += "#{(hash_charge_data[:qtd] - 1)}".pad_with_left_zeroes! 11 # [O] Z 11
    line += "#{counter}".pad_with_left_zeroes! 11 # [P] Z 11
    line += "".pad_with_right_spaces! 552 # [Q] C 553
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content }
end

# Reads the data in the "Detalhe - Tabela de CARGO  FUNÇ" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def job_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    if seq == 0 # jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[D] Z 03 means that data in column D must be of only-numeric type and
    #       having 3 characters
    line = ""
    line += "TC".pad_with_right_spaces! 2 # [A] C 02
    line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "".pad_with_right_spaces! 33 # [E] C 33
    line += "#{row[5]}".pad_with_left_zeroes! 7 # [F] Z 07
    line += "#{row[6]}".pad_with_right_spaces! 30 # [G] C 30
    line += "".pad_with_right_spaces! 665 # [H] C 666
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content, :qtd => seq }
end

# Reads the data in the "Detalhe - Tabela de VENCIMENTO" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def maturities_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    if seq == 0 # Jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[J] Z 02 means that data in column J must be of only-numeric type and
    #       having 2 characters
    line = ""
    line += "TV".pad_with_right_spaces! 2 # [A] C 02
    line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "".pad_with_right_spaces! 33 # [E] C 33
    line += "#{row[5]}".pad_with_left_zeroes! 6 # [F] Z 06
    line += "#{row[6]}".pad_with_right_spaces! 30 # [G] C 30
    line += "#{row[7]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [H] C 01
    line += "#{row[8]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [I] C 01
    line += "#{row[9]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [J] Z 02
    line += "".pad_with_right_spaces! 662 # [K] C 663
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content, :qtd => seq }
end

# Reads the data in the "Detalhe - Dados PESSOAIS  FUNCI" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line, the records number
# and total salaries, discounts, and liquids
def personal_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter
  total_gross = 0
  total_net = 0
  total_discounts = 0

  sheet.rows.each do |row|
    if seq == 0 # jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[BG] Z 05 means that data in column BG must be of only-numeric type and
    #        having 5 characters
    line = ""
    line += "F1".pad_with_right_spaces! 2 # [A] C 02
    line += "".pad_with_right_spaces! 7 # [B] C 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "#{row[4]}".pad_with_left_zeroes! 3 # [E] Z 03
    line += "#{row[5]}".pad_with_left_zeroes! 7 # [F] Z 07
    line += "".pad_with_right_spaces! 5 # [G] C 05
    line += "#{row[7]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [H] Z 02
    line += "#{row[8]}".split(' ')[0].to_s.pad_with_left_zeroes! 4 # [I] Z 04
    line += "#{row[9]}".pad_with_left_zeroes! 10 # [J] Z 10
    line += "#{row[10]}".pad_with_left_zeroes! 2 # [K] Z 02
    line += "#{row[11]}".pad_with_left_zeroes! 2 # [L] Z 02
    line += "#{row[12]}".pad_with_left_zeroes! 8 # [M] Z 08
    line += "#{row[13]}".pad_with_right_spaces! 30 # [N] C 30
    line += "#{row[14]}".pad_with_left_zeroes! 8 # [O] Z 08
    line += "#{row[15]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [P] C 01
    line += "#{row[16]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [Q] C 01
    line += "#{row[17]}".split(' ')[0].to_s.pad_with_right_spaces! 2 # [R] C 02
    line += "#{row[18]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [S] Z 02
    line += "#{row[19]}".pad_with_right_spaces! 30 # [T] C 30
    line += "#{row[20]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [U] C 01
    line += "#{row[21]}".split(' ')[0].to_s.pad_with_left_zeroes! 1 # [V] Z 01
    line += "#{row[22]}".pad_with_left_zeroes! 11 # [W] C 11
    line += "#{row[23]}".split(' ')[0].to_s.pad_with_right_spaces! 2 # [X] C 02
    line += "#{row[24]}".pad_with_left_zeroes! 9 # [Y] Z 09
    line += "#{row[25]}".pad_with_left_zeroes! 2 # [Z] Z 02
    line += "#{row[26]}".pad_with_left_zeroes! 11 # [AA] Z 11
    line += "#{row[27]}".pad_with_left_zeroes! 7 # [AB] Z 07
    line += "#{row[28]}".pad_with_left_zeroes! 5 # [AC] Z 05
    line += "#{row[29]}".split(' ')[0].to_s.pad_with_right_spaces! 2 # [AD] C 02
    line += "#{row[30]}".pad_with_left_zeroes! 11 # [AE] Z 11
    line += "#{row[31]}".pad_with_left_zeroes! 3 # [AF] Z 03
    line += "#{row[32]}".pad_with_left_zeroes! 4 # [AG] Z 04
    line += "#{row[33]}".pad_with_right_spaces! 30 # [AH] C 30
    line += "#{row[34]}".split(' ')[0].to_s.pad_with_right_spaces! 2 # [AI] C 02
    line += "#{row[35]}".pad_with_right_spaces! 30 # [AJ] C 30
    line += "#{row[36]}".pad_with_left_zeroes! 8 # [AK] Z 08
    line += "#{row[37]}".pad_with_right_spaces! 20 # [AL] C 20
    line += "#{row[38]}".pad_with_left_zeroes! 3 # [AM] Z 03
    line += "#{row[39]}".pad_with_left_zeroes! 3 # [AN] Z 03
    line += "#{row[40]}".pad_with_left_zeroes! 8 # [AO] Z 08
    line += "#{row[41]}".pad_with_left_zeroes! 4 # [AP] Z 04
    line += "#{row[42]}".pad_with_left_zeroes! 3 # [AQ] Z 03
    line += "#{row[43]}".pad_with_left_zeroes! 8 # [AR] Z 08
    line += "#{row[44]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [AS] C 01
    line += "#{row[45]}".pad_with_left_zeroes! 7 # [AT] Z 07
    line += "#{row[46]}".pad_with_left_zeroes! 5 # [AU] Z 05
    line += "#{row[47]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [AV] C 01
    line += "#{row[48]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [AW] C 01
    line += "#{row[49]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [AX] Z 02
    line += "#{row[50]}".split(' ')[0].to_s.pad_with_left_zeroes! 3 # [AY] Z 03
    line += "#{row[51]}".pad_with_left_zeroes! 7 # [AZ] Z 07
    line += "#{row[52]}".pad_with_left_zeroes! 8 # [BA] Z 08
    line += "#{row[53]}".pad_with_left_zeroes! 2 # [BB] Z 02
    line += "#{row[54]}".pad_with_left_zeroes! 5 # [BC] Z 05
    line += "#{row[55]}".pad_with_left_zeroes! 5 # [BD] Z 05
    line += "#{row[56]}".pad_with_left_zeroes! 5 # [BE] Z 05
    line += "#{row[57]}".pad_with_left_zeroes! 5 # [BF] Z 05
    line += "#{row[58]}".pad_with_left_zeroes! 5 # [BG] Z 05
    line += "#{row[59]}".split(' ')[0].to_s.pad_with_left_zeroes! 2 # [BH] Z 02
    line += "#{row[60]}".split(' ')[0].to_s.pad_with_left_zeroes! 3 # [BI] Z 03
    line += "#{row[61]}".pad_with_left_zeroes! 8 # [BJ] Z 08
    line += "#{row[62]}".pad_with_left_zeroes! 3 # [BK] Z 03
    line += "#{row[63]}".pad_with_left_zeroes! 3 # [BL] Z 03
    line += "#{row[64]}".pad_with_left_zeroes! 2 # [BM] Z 02
    line += "#{row[65]}".pad_with_right_spaces! 10 # [BN] C 10
    line += "#{row[66]}".pad_with_right_spaces! 1 # [BO] C 01
    line += "#{row[67]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [BP] C 01
    line += "#{row[68]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [BQ] C 01
    line += "#{row[69]}".pad_with_left_zeroes! 2 # [BR] Z 02
    line += "#{row[70]}".pad_with_left_zeroes! 2 # [BS] Z 02
    line += "#{row[71]}".pad_with_left_zeroes! 4 # [BT] Z 04
    line += "#{row[72]}".pad_with_left_zeroes! 5 # [BU] Z 05
    line += "#{row[73]}".pad_with_left_zeroes! 2 # [BV] Z 02
    line += "#{row[74]}".pad_with_left_zeroes! 5 # [BW] Z 05
    line += "#{row[75]}".pad_with_left_zeroes! 5 # [BX] Z 05
    line += "#{row[76]}".pad_with_right_spaces! 1 # [BY] C 01
    line += "#{row[77]}".pad_with_right_spaces! 1 # [BZ] C 01
    line += "#{row[78]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [CA] C 01
    line += "#{row[79]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [CB] C 01
    line += "#{row[80]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [CC] C 01
    line += "#{row[81]}".pad_with_left_zeroes! 2 # [CD] Z 02
    line += "#{row[82]}".pad_with_left_zeroes! 2 # [CE] Z 02
    line += "#{row[83]}".pad_with_left_zeroes! 7 # [CF] Z 07
    line += "#{row[84]}".pad_with_left_zeroes! 3 # [CG] Z 03
    line += "#{row[85]}".pad_with_left_zeroes! 3 # [CH] Z 03
    line += "#{row[86]}".pad_with_left_zeroes! 3 # [CI] Z 03
    line += "#{row[87]}".pad_with_left_zeroes! 5 # [CJ] Z 05
    line += "#{row[88]}".pad_with_left_zeroes! 5 # [CK] Z 05
    line += "#{row[89]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [CL] C 01
    line += "#{row[90]}".pad_with_left_zeroes! 13 # [CM] Z 13
    line += "#{row[91]}".pad_with_left_zeroes! 7 # [CN] Z 07
    line += "#{row[92]}".pad_with_left_zeroes! 7 # [CO] Z 07
    line += "#{row[93]}".pad_with_left_zeroes! 7 # [CP] Z 07
    line += "#{row[94]}".pad_with_left_zeroes! 7 # [CQ] Z 07
    line += "#{row[95]}".pad_with_left_zeroes! 7 # [CR] Z 07
    line += "#{row[96]}".pad_with_left_zeroes! 7 # [CS] Z 07
    line += "#{row[97]}".pad_with_left_zeroes! 3 # [CT] Z 03
    line += "#{row[98]}".pad_with_left_zeroes! 3 # [CU] Z 03
    line += "#{row[99]}".pad_with_left_zeroes! 2 # [CV] Z 02
    line += "#{row[100]}".pad_with_right_spaces! 10 # [CW] C 10
    line += "#{row[101]}".pad_with_right_spaces! 1 # [CX] C 01
    line += "#{row[102]}".pad_with_left_zeroes! 7 # [CY] Z 07
    line += "#{row[103]}".pad_with_left_zeroes! 3 # [CZ] Z 03
    line += "#{row[104]}".pad_with_left_zeroes! 3 # [DA] Z 03
    line += "#{row[105]}".pad_with_left_zeroes! 2 # [DB] Z 02
    line += "#{row[106]}".pad_with_right_spaces! 10 # [DC] C 10
    line += "#{row[107]}".pad_with_right_spaces! 1 # [DD] C 01
    line += "#{row[108]}".pad_with_left_zeroes! 7 # [DE] Z 07
    line += "#{row[109]}".pad_with_left_zeroes! 3 # [DF] Z 03
    line += "#{row[110]}".pad_with_left_zeroes! 3 # [DG] Z 03
    line += "#{row[111]}".pad_with_left_zeroes! 2 # [DH] Z 02
    line += "#{row[112]}".pad_with_right_spaces! 10 # [DI] C 10
    line += "#{row[113]}".pad_with_right_spaces! 1 # [DJ] C 01
    line += "#{row[114]}".pad_with_left_zeroes! 7 # [DK] Z 07
    line += "#{row[115]}".pad_with_left_zeroes! 3 # [DL] Z 03
    line += "#{row[116]}".pad_with_left_zeroes! 3 # [DM] Z 03
    line += "#{row[117]}".pad_with_left_zeroes! 2 # [DN] Z 02
    line += "#{row[118]}".pad_with_right_spaces! 10 # [DO] C 10
    line += "#{row[119]}".pad_with_right_spaces! 1 # [DP] C 01
    line += "#{row[120]}".pad_with_left_zeroes! 7 # [DQ] Z 07
    line += "#{row[121]}".pad_with_left_zeroes! 3 # [DR] Z 03
    line += "#{row[122]}".pad_with_left_zeroes! 3 # [DS] Z 03
    line += "#{row[123]}".pad_with_left_zeroes! 2 # [DT] Z 02
    line += "#{row[124]}".pad_with_right_spaces! 10 # [DU] C 10
    line += "#{row[125]}".pad_with_right_spaces! 1 # [DV] C 01
    line += "#{row[126]}".split(' ')[0].to_s.pad_with_left_zeroes! 3 # [DW] Z 03
    line += "#{row[127]}".pad_with_left_zeroes! 8 # [DX] Z 08
    line += "#{row[128]}".pad_with_left_zeroes! 5 # [DY] Z 05
    line += "#{row[129]}".pad_with_left_zeroes! 5 # [DZ] Z 05
    line += "#{row[130]}".split(' ')[0].to_s.pad_with_left_zeroes! 3 # [EA] Z 03
    # total gross
    line += "#{row[131]}".pad_with_left_zeroes! 13 # [EB] Z 13

    total_gross += row[131].to_s.to_i
    # total discounts
    line += "#{row[132]}".pad_with_left_zeroes! 13 # [EC] Z 13

    total_discounts += row[132].to_s.to_i
    # total net
    line += "#{row[133]}".pad_with_left_zeroes! 13 # [ED] Z 13

    total_net += row[133].to_s.to_i
    line += "#{row[134]}".split(' ')[0].to_s.pad_with_right_spaces! 2 # [EE] C 02
    line += "#{row[135]}".pad_with_right_spaces! 1 # [EF] C 01
    line += "#{row[136]}".pad_with_left_zeroes! 7 # [EG] Z 07
    line += "#{row[137]}".pad_with_left_zeroes! 15 # [EH] Z 15
    line += "#{row[138]}".pad_with_left_zeroes! 3 # [EI] Z 03
    line += "#{row[139]}".split(' ')[0].to_s.pad_with_left_zeroes! 1 # [EJ] Z 01
    line += "#{row[140]}".split(' ')[0].to_s.pad_with_right_spaces! 1 # [EK] C 01
    line += "".pad_with_right_spaces! 2 # [EL] C 03
    line += "\n" # considered as one char

    content += line
    seq += 1
  end


  { :content => content,
    :qtd => seq,
    :total_gross => total_gross,
    :total_net => total_net,
    :total_discounts => total_discounts }
end

# Reads the data in the "Detalhe - Tabela de FAIXA SALAR" tab of worksheet and
# then converts them to lines of text on the format expected.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def salary_range_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    if seq == 0	# jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[N] Z 05 means that data in column N must be of only-numeric type and
    #       having 5 characters
    line = ""
    line += "TF".pad_with_right_spaces! 2 # [A] C 02
    line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "".pad_with_right_spaces! 33 # [E] C 33
    line += "#{row[5]}".pad_with_right_spaces! 10 # [F] C 10
    line += "#{row[6]}".pad_with_left_zeroes! 13 # [G] Z 13
    line += "#{row[7]}".pad_with_left_zeroes! 5 # [H] Z 05
    line += "#{row[8]}".pad_with_left_zeroes! 13 # [I] Z 13
    line += "#{row[9]}".pad_with_left_zeroes! 5 # [J] Z 05
    line += "#{row[10]}".pad_with_left_zeroes! 13 # [K] Z 13
    line += "#{row[11]}".pad_with_left_zeroes! 5 # [L] Z 05
    line += "#{row[12]}".pad_with_left_zeroes! 13 # [M] Z 13
    line += "#{row[13]}".pad_with_left_zeroes! 5 # [N] Z 05
    line += "".pad_with_right_spaces! 620 # [O] C 621
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content, :qtd => seq }
end

# Reads the data in the "Detalhe - Tabela de UNIDADES AD" tab of worksheet and
# then converts them to lines of text on the format expected by PRODESP.
#
# @param [<#SXR::Sheet>] sheet Data inside this tab
# @return [Hash] Content of the tab converted to text line and the records number
def unit_data sheet
  content = '' # content of text line
  seq = 0 # sequence counter

  sheet.rows.each do |row|
    if seq == 0 # jumps the header
      seq += 1
      next
    end

    break if "#{row[2]}".empty? # execute only while the "Órgão" field is filled

    # NOTATION: [SHEET COLUMN] 	VALUE TYPE('C' for alphanumeric or 'Z' for only
    #                                      numeric)	LIMIT SIZE
    # Example:
    # 	[H] C 70 means that data in column H must be of alphanumeric type and
    #       having 70 characters
    line = ""
    line += "TU".pad_with_right_spaces! 2 # [A] C 02
    line += "#{seq}".pad_with_left_zeroes! 7 # [B] Z 07
    line += "#{row[2]}".pad_with_left_zeroes! 2 # [C] Z 02
    line += "#{row[3]}".pad_with_left_zeroes! 3 # [D] Z 03
    line += "".pad_with_right_spaces! 33 # [E] C 33
    line += "#{row[5]}".pad_with_left_zeroes! 3 # [F] Z 03
    line += "#{row[6]}".pad_with_left_zeroes! 7 # [G] Z 07
    line += "#{row[7]}".pad_with_right_spaces! 70 # [H] C 70
    line += "#{row[8]}".pad_with_left_zeroes! 15 # [I] Z 15
    line += "#{row[9]}".pad_with_right_spaces! 70 # [J] C 70
    line += "".pad_with_right_spaces! 537 # [K] C 538
    line += "\n" # considered as one char

    content += line
    seq += 1
  end

  { :content => content, :qtd => seq}
end

# Writes the lines proccessed previously in the text file to be sent to PRODESP.
#
#  @param [Hash] file_path Path of file to be generated
#  @param [Hash] hash_header Hash that contains all the data of "Header" tab
#  @param [Hash] hash_charge_data Hash that contains all the data of
#                                 "Detalhe - ENCARGOS SOCIAIS  BEN" tab
#  @param [Hash] hash_personal_data Hash that contains all the data of
#                                   "Detalhe - Dados PESSOAIS  FUNCI" tab
#  @param [Hash] hash_financial_data Hash that contains all the data of
#                                    "Detalhe - Dados FINANCEIROS" tab
#  @param [Hash] hash_job_data Hash that contains all the data of
#                              "Detalhe - Tabela de CARGO  FUNÇ" tab
#  @param [Hash] hash_maturities_data Hash that contains all the data of
#                                     "Detalhe - Tabela de VENCIMENTO" tab
#  @param [Hash] hash_salary_range_data Hash that contains all the data of
#                                       "Detalhe - Tabela de FAIXA SALAR" tab
#  @param [Hash] hash_unit_data Hash that contains all the data of
#                               "Detalhe - Tabela de UNIDADES AD" tab
def write_file(file_path, hash_header, hash_charge_data, hash_personal_data,
               hash_financial_data, hash_job_data, hash_maturities_data,
               hash_salary_range_data, hash_unit_data)
  File.open(file_path, 'w') do |f|
    f.write hash_header[:content]
    f.write hash_charge_data[:content]
    f.write hash_personal_data[:content]
    f.write hash_financial_data[:content]
    f.write hash_job_data[:content]
    f.write hash_maturities_data[:content]
    f.write hash_salary_range_data[:content]
    f.write hash_unit_data[:content]
  end
end
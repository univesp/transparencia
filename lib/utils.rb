#!/bin/env ruby
# encoding: utf-8
# This script extends the String class to provide padding to the values of fields
# in order to fit them the constraints of file to be sent to PRODESP.

class String
  # Fills a text with spaces to the right until the maximum size allowed for the field.
  # Often used in alphanumeric fields.
  #
  # @param [Hash] size Maximum size of text to be padded
  # @param [Hash] padstr Character to be used to pad the text. Default is ' '
  def pad_with_right_spaces!(size, padstr=' ')
    self[0..size-1].to_s.ljust size, padstr
  end

  # Fills a text with zeros to the left until the maximum size allowed for the field.
  # Often used in only-numeric fields.
  #
  # @param [Hash] size Maximum size of text to be padded
  # @param [Hash] padstr Character to be used to pad the text. Default is '0'
  def pad_with_left_zeroes!(size, padstr='0')
    str = self[0..size-1]
    has_letters = str.match /[A-Z]/i
    if has_letters # do not convert to integer
      str.rjust size, padstr
    else # convert to integer
      str.to_i.to_s.rjust size, padstr
    end
  end
end
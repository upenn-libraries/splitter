#!/usr/bin/env ruby

require 'csv'
require 'rubyXL'
require 'logger'

def missing_args?
  return (ARGV[0].nil? || ARGV[1].nil? || ARGV[2].nil?)
end

def set_up_spreadsheet(workbook, headers)
  worksheet = workbook.worksheets[0]
  set_headers(worksheet, headers)
end

def set_headers(worksheet, headers)
  headers.each_with_index do |header, i|
    worksheet.add_cell(0,i, header)
  end
end

def identifier_missing?(identifier_location)
  return identifier_location.nil? || identifier_location.empty?
end

def offset(index)
  return index == 0 ? index : index-1
end

def filename(filename_string)
  return filename_string.downcase.gsub(/[^0-9A-Za-z.\-]/, '_')
end

def extract_rows(filename)
  contents_array = []

  if File.extname(filename).downcase == '.csv'
    csv = filename
    headers = CSV.readlines(csv, :encoding => 'ISO8859-1:utf-8').first
    headers.freeze
    contents_array = CSV.read(csv, :encoding => 'ISO8859-1:utf-8')
  elsif File.extname(filename).downcase == '.xlsx'
    xlsx = RubyXL::Parser.parse(filename)
    worksheet = xlsx[0]
    contents_array = []
    worksheet.sheet_data.rows.each do |row|
      row_array = []
      row.cells.each do |cell|
        value = cell.nil? ? '' : cell.value
        row_array << value
      end
      contents_array << row_array
    end
    headers = contents_array.first
  else
    abort("Unsupported file extension: #{File.extname(filename).downcase}")
  end

  contents_array.shift

  return headers, contents_array
end

abort('Specify a path to an Excel or CSV file, an integer for column whose value should be considered the identifier for each row, and the batch size that should be used for this job') if missing_args?

progname = File.basename __FILE__
logger = Logger.new('| tee logger.log')
logger.level = Logger::INFO
logger.info(progname) { 'Script run started' }

warning_logger = Logger.new('| tee warning_logger.log')
warning_logger.level = Logger::WARN

headers, content_array = extract_rows(ARGV[0])

index_of_identifier = offset(ARGV[1].to_i)
batch_size = ARGV[2].to_i

content_array.each_slice(batch_size).with_index do |batch, batch_index|
  workbook = RubyXL::Workbook.new
  set_up_spreadsheet(workbook, headers)
  batch.each_with_index do |row, index|
    if identifier_missing?(row[index_of_identifier])
      logger.warn(progname) { "missing identifier for row ##{index}, logging and skipping..." }
      warning_logger.warn(progname) { "Row ##{index} in #{csv} missing identifier specified as being in column #{index_of_identifier}" }
      next
    end
    row.each_with_index do |value, i|
      workbook[0].add_cell(index+1,i,value)
    end
  end
  spreadsheet_name = "#{filename("#{batch.first[index_of_identifier]}_through_#{batch.last[index_of_identifier]}")}.xlsx"
  logger.info(progname) { "Writing spreadsheet #{spreadsheet_name}..." }
  workbook.write(spreadsheet_name)
end

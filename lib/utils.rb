require 'active_support'
require 'active_support/inflector'
require 'colorize'
# require 'chronic'
require 'dotenv/tasks'
require 'dotenv'
require 'erb'

# Login to Google with a saved session and set spreadsheet
def login
  system('clear')
  puts 'Authorizing...'.green

  @session ||= GoogleDrive.saved_session('config.json')
  @ws ||= @session.spreadsheet_by_key(ENV['SPREADSHEET_KEY']).worksheets[0]
end

def set_headers
  @headers ||= {}
  login
  counter = 1
  (1..@ws.num_cols).each do |col|
    @headers[@ws[1, col].gsub(/\s+/, '_').downcase.to_sym] = counter
    counter += 1
  end

  @headers
end

def smart_add_url_protocol(url)
  unless url[/\Ahttp:\/\//] || url[/\Ahttps:\/\//] || url.empty?
    url = "http://#{url}"
  end
  url
end

def project_hash(row)
  @headers ||= set_headers

  project = {
    title: @ws[row, @headers[:title]],
    year: @ws[row, @headers[:year]],
    end_range: @ws[row, @headers[:end_range]],
    formats: @ws[row, @headers[:formats]].to_json,
    geographic_extant: @ws[row, @headers[:geographic_extant]].to_json,
    institution: @ws[row, @headers[:institution]].to_json,
    institution_type: @ws[row, @headers[:institution_type]].to_json,
    primary_contact: @ws[row, @headers[:primary_contact]].to_json,
    size: @ws[row, @headers[:size]].to_json,
    start_range: @ws[row, @headers[:start_range]].to_json,
    summary: @ws[row, @headers[:summary]].to_json,
    hc_id: @ws[row, @headers[:hc_id]].to_json,
    program: @ws[row, @headers[:program]].to_json,
    grant_type: @ws[row, @headers[:grant_type]].to_json,
    project_url: @ws[row, @headers[:project_url]].to_json,
    address1:  @ws[row, @headers[:address1]].to_json,
    address2:  @ws[row, @headers[:address2]].to_json,
    city:  @ws[row, @headers[:city]].to_json,
    state:  @ws[row, @headers[:state]].to_json,
    zip:  @ws[row, @headers[:zip]].to_json,
    repository_url:  @ws[row, @headers[:repository_url]].to_json,
    funded:  @ws[row, @headers[:funded]].to_json,
    p1_institution:  @ws[row, @headers[:p1_institution]].to_json,
    p2_name:  @ws[row, @headers[:p2_name]].to_json,
    p2_institution:  @ws[row, @headers[:p2_institution]].to_json,
    p3_name:  @ws[row, @headers[:p3_name]].to_json,
    p3_institution:  @ws[row, @headers[:p3_institution]].to_json,
    material_string: @ws[row, @headers[:material_string]].to_json,
    collaborating_institution:  @ws[row, @headers[:collaborating_institution]].to_json,
    grant_amount:  @ws[row, @headers[:grant_amount]].to_json,
    material_description:  @ws[row, @headers[:material_description]].to_json
  }
  project[:file_path] = filename(project)
  project
end

def filename(project)
  project_name = ActiveSupport::Inflector.parameterize(project[:title])
  "_projects/#{project[:year]}-#{project_name.slice(0..50)}.md"
end

def render_erb(template_path)
  template = File.open(template_path, 'r').read
  erb = ERB.new(template)
  erb.result(binding)
end

def spreadsheet
  @ws ||= @session.spreadsheet_by_key(ENV['SPREADSHEET_KEY'])
end

def write_file(path, contents)
  file = File.open(path, 'w')
  file.write(contents)
rescue IOError => error
  puts 'File not writable. Check your permissions'
  puts error.inspect
ensure
  file.close unless file.nil?
end

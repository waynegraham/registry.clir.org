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
    formats: @ws[row, @headers[:formats]],
    geographic_extant: @ws[row, @headers[:geographic_extant]],
    institution: @ws[row, @headers[:institution]],
    institution_type: @ws[row, @headers[:institution_type]],
    primary_contact: @ws[row, @headers[:primary_contact]],
    size: @ws[row, @headers[:size]],
    start_range: @ws[row, @headers[:start_range]],
    summary: @ws[row, @headers[:summary]],
    hc_id: @ws[row, @headers[:hc_id]],
    program: @ws[row, @headers[:program]],
    grant_type: @ws[row, @headers[:grant_type]],
    project_url: @ws[row, @headers[:project_url]],
    address1:  @ws[row, @headers[:address1]],
    address2:  @ws[row, @headers[:address2]],
    city:  @ws[row, @headers[:city]],
    state:  @ws[row, @headers[:state]],
    zip:  @ws[row, @headers[:zip]],
    repository_url:  @ws[row, @headers[:repository_url]],
    funded:  @ws[row, @headers[:funded]],
    p1_institution:  @ws[row, @headers[:p1_institution]],
    p2_name:  @ws[row, @headers[:p2_name]],
    p2_institution:  @ws[row, @headers[:p2_institution]],
    p3_name:  @ws[row, @headers[:p3_name]],
    p3_institution:  @ws[row, @headers[:p3_institution]],
    material_string: @ws[row, @headers[:material_string]],
    collaborating_institution:  @ws[row, @headers[:collaborating_institution]],
    grant_amount:  @ws[row, @headers[:grant_amount]],
    material_description:  @ws[row, @headers[:material_description]]
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

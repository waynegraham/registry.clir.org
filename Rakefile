require 'active_support'
require 'active_support/inflector'
require 'colorize'
require 'dotenv'
require 'dotenv/tasks'
require 'erb'
require 'google_drive'
require 'ra11y'

require_relative 'lib/utils.rb'

Dotenv.load

task default: 'import:all'

desc 'Clean the `_project` directory'
task :clean do
  FileUtils.rm_f Dir.glob('_projects/**/*')
end

namespace :import do
  desc 'Import all projects'
  task all: %i[projects]

  desc 'Import projects from the Google Spreadsheet'
  task :projects do
    login

    (2..@ws.num_rows).each do |row|
      @project = project_hash(row)
      contents = render_erb('templates/project.html.erb')
      file_path = @project[:file_path]
      write_file(file_path, contents)
      puts "Writing the project for #{@project[:title]}".green
    end

  end
end

require 'capybara'
require 'capybara/rspec'
require 'capybara/dsl'

Capybara.register_driver :selenium_ie do |app|
	Capybara::Selenium::Driver.new(app, :browser => :ie)
end
Capybara.run_server = false
Capybara.current_driver = :selenium_ie
Capybara.app_host = 'https://www.google.com.sg/?gfe_rd=cr&dcr=0&ei=JkG2WYqII4OEoAPyoLbgCw'
Capybara.default_max_wait_time = 5

module CapybaraTest
	class Test
		include Capybara::DSL
		def initialize
			Capybara.default_driver = :selenium_ie
		end
		def test_google
			visit('/')
			visit('/')
			visit('https://www.microsoft.com/en-sg/')
			visit('https://bing.com')
			visit('/')
			visit('/')
			sleep(10)
			puts "yahoo!"
		end
		def test_work
			visit('https://yahoo.com.sg')
			puts "whipee!"
		end
	end
end

a = CapybaraTest::Test.new
a.test_google
a.test_work
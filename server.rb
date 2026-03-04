require 'webrick'

server = WEBrick::HTTPServer.new(
  Port: 8080,
  DocumentRoot: File.expand_path(File.dirname(__FILE__)),
  BindAddress: '127.0.0.1'
)

trap('INT') { server.shutdown }
trap('TERM') { server.shutdown }

server.start

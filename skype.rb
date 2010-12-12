require "rubygems"
require "win32ole"
require "drb/drb"
require "kconv"

$KCODE="u"
class WIN32OLE
  def sort
    ole_methods.join("\t").split("\t").sort
  end

  def to_a
    ary=[]
    each do |o|
      ary.push o
    end
    ary
  end
end

class Skype
  attr_accessor :skype, :events, :cur_chat, :cur_chats
  def initialize
    @skype=WIN32OLE.new "Skype4COM.Skype"
    #@events=WIN32OLE_EVENT.new(@skype, "_ISkypeEvents")
    @cur_chat = nil
    @cur_chat = []
  end

  def puts(str)
    Kernel.puts(Kconv.toutf8(str))
  end

  def get_chats
    @skype.Chats.to_a.map{|c| Chat.new(c)}
  end
  alias :chats :get_chats

  def get_recent_chats
    @skype.RecentChats.to_a.map{|c| Chat.new(c)}
  end
  alias :recent_chats :get_recent_chats
  alias :rc :get_recent_chats

  #チャットのFriendlyNameだけを表示
  def get_head
    Kernel.puts "sending now..."
    @cur_chats = get_recent_chats
    @cur_chats.each_index do |index|
      chat = @cur_chats[index]
      message = chat.messages[0]
      if(!message.has_read?)
        puts chat.name
      else
        break
      end
    end
    nil
  end
  alias :head :get_head
  alias :h :get_head

  #既読にする？意味ない？
#  def read
#    recent_chats.each do |chat|
#      #chat.OpenWindow
#      chat.ClearRecentMessages
#    end
#    nil
#  end
#  alias :r :read

  #Skypeのウィンドウを最前面に
  def open
    chats[0].OpenWindow
  end
  alias :o :open

  #TODO:チャットモードを切り替える 空の時はsendメソッドは使えない
  def set_chat(index)
    @cur_chat = @cur_chats[index]
  end
  alias :set :set_chat

  def show_chats
    @cur_chats = get_chats
    show_chats_with_number
  end
  alias :sc :show_chats

  def show_recent_chats
    @cur_chats = get_recent_chats
    show_chats_with_number
  end
  alias :src :show_recent_chats
  alias :show :show_recent_chats

  def show_chats_with_number
    @cur_chats.each_index do |i|
      puts "#{i}  #{@cur_chats[i].name}"
    end
    nil
  end

  def send(str)
    @cur_chat.send(str)
    nil
  end

  def read_cur_chats
    Kernel.puts "sending now..."
    chats = cur_chats || get_recent_chats
    chats.each do |chat|
      chat.read
    end
    nil
  end
  alias :rcc :read_cur_chats

  class Chat
    def initialize(chat)
      @chat = chat
    end

    def name
      begin
        @chat.FriendlyName
      rescue
        ""
      end
    end

    def send(str)
      @chat.SendMessage(str)
    end

    def read
      messages.reverse.each do |message|
        if(!message.has_read? && message.is_normal_message?)
          name = message.FromDisplayName
          edited = message.EditedBy
          puts Kconv.toutf8 "#{@chat.Name} #{message.Type} #{message.Role} #{edited} #{@chat.name} #{message.Status} #{name} #{message.Body} #{message.TimeStamp}"
        end
      end
      nil
    end
    alias :r :read

    def messages
      @chat.Messages.to_a.map{|m| Message.new(m)}
    end

    def recent_messages
      @chat.RecentMessages.to_a.map{|m| Message.new(m)}
    end

    def method_missing(meth, *args)
      if @chat.ole_methods.find{|m| meth.to_s==m.to_s}
        @chat.__send__(meth, *args)
      else
        send "#{meth} #{args.join(',')}"
      end
    end

    class Message
      STATUS_READ = 2
      TYPE_NORMAL_MESSAGE = 4
      def initialize(message)
        @message = message
      end

      def has_read?
        @message.Status != STATUS_READ
      end

      def is_normal_message?
        @message.Type == TYPE_NORMAL_MESSAGE
      end

      def method_missing(meth, *args)
        @message.__send__(meth, *args)
      end
    end
  end
end

class SkypeShell
  def initialize
    @skype = nil
    @chat = nil
  end

  def start
    loop do
      begin
        print("#{(@chat ? Kconv.toutf8(@chat.name) : "")}> ")
      rescue => e
        puts e
        print("> ")
      end
      input = gets.chomp
      break if input == "quit"
      if(input.size.zero?)
        puts ""
      else
        case input
        when "get_head", "head", "h", "show_recent", "sr", "show"
          @skype = Skype.new
        else
          @skype = Skype.new unless @skype
        end
        #TODO:eval使わない形式に直す
        if(@chat && /^(send .+|read|r)$/ =~ input)
          begin
            result = eval("@chat.#{Kconv.tosjis(input)}")
          rescue => e
            result = e
          end
        else
          begin
            result = eval("@skype.#{Kconv.tosjis(input)}")
          rescue => e
            result = e
          end
        end
        if(/^set(_chat)? .+$/ =~ input)
          @chat = result
        end
        puts(result)
      end
    end
  end
end
Skysh.new.start


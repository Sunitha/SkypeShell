require "rubygems"
require "win32ole"
require "kconv"

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
  def initialize(char_code="u")
    @skype=WIN32OLE.new "Skype4COM.Skype"
    @skype.Attach()
    #@events=WIN32OLE_EVENT.new(@skype, "_ISkypeEvents")
    @cur_chat = nil
    @cur_chat = []
    case char_code
    when "s", "sjis"
      @kconv_method = "tosjis"
    else
      @kconv_method = "toutf8"
    end
  end

  def quit
    exit
  end

  def call(handle)
    @skype.PlaceCall(handle)
  end

  def get_active_calls
    @skype.ActiveCalls.to_a
  end

  def puts(str)
    begin
      Kernel.puts(Kconv.__send__(@kconv_method, str)) unless str.nil?
    rescue
      Kernel.puts ""
    end
  end

  def print(str)
    begin
      Kernel.print(Kconv.__send__(@kconv_method, str)) unless str.nil?
    rescue
      Kernel.print ""
    end
  end

  def get_chats
    @skype.Chats.to_a.map{|c| Chat.new(self, c)}
  end
  alias :chats :get_chats

  def get_recent_chats
    @skype.RecentChats.to_a.map{|c| Chat.new(self, c)}
  end
  alias :recent_chats :get_recent_chats
  alias :rc :get_recent_chats

  #チャットのFriendlyNameだけを表示
  def get_head
    puts "sending now..."
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

  #Skypeのウィンドウを最前面に
  def open
    (cur_chat || get_recent_chats[0]).OpenWindow
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
    attr_reader :chat
    def initialize(skype, chat)
      @skype = skype
      @chat = chat
      @cur_call = nil
    end

    def name
      begin
        @chat.FriendlyName
      rescue
        ""
      end
    end

    def call
      handles = get_handles
      handles.each do |handle|
        begin
          @cur_call = Call.new(@skype.call(handle))
          break
        rescue
        end
      end
      handles[1..(handles.size-1)].each do |handle|
        begin
          @cur_call.join(handle)
        rescue
        end
      end
      @cur_call
    end
    alias :c :call

    def finish
      begin
        @cur_call.finish
      rescue
        ""
      end
    end
    alias :f :finish

    def get_handles
      @chat.Members.to_a.map{|m| m.Handle}
    end

    def show_members
      @chat.Members.to_a.each do |m|
        #@skype.puts m.DisplayName
        @skype.puts m.Handle
      end
      nil
    end
    alias :members :show_members
    alias :m :show_members

    def send(str)
      @chat.SendMessage(str)
    end

    def read
      puts messages.size
      messages.reverse.each do |message|
        if(!message.has_read? && message.is_normal_message?)
          @skype.puts "#{@chat.Name} #{message.Type} #{message.Role} #{edited} #{@chat.name} #{message.Status} #{name} #{message.Body} #{message.TimeStamp}"
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

  class Call
    attr_reader :call
    def initialize(call)
      @call = call
    end

    def finish
      @call.Finish
    end
  end
end

class SkypeShell
  def initialize(char_code="u")
    @skype = nil
    @chat = nil
    @char_code = char_code
  end

  def start
    renew_skype
    loop do
      puts_prompt
      input = gets.chomp
      if(input.size.zero?)
        @skype.puts ""
      else
        renew_skype(input)
        result = eval_input(input)
        #renew_chat(input, result)
        #@skype.puts(result)
        #@skype.puts ""
      end
    end
  end

  def puts_prompt
    begin
      @skype.print("#{(@chat ? @chat.name : "")}> ")
    rescue => e
      puts e
      @skype.print("> ")
    end
  end

  #TODO:eval使わない形式に直す
  def eval_input(input)
    if(/^:(.*)$/ =~ input)
      begin
        result = @chat.send(Kconv.tosjis($1))
      rescue => e
        result = e
      end
    elsif(@chat && /^(send .+|read|r|call|c|finish|f|members|m)$/ =~ input)
      begin
        result = eval("@chat.#{encode(input)}")
      rescue => e
        result = e
      end
    else
      begin
        result = eval("@skype.#{encode(input)}")
      rescue => e
        result = e
      end
    end
    result
  end

  def encode(str)
    if(str)
      Kconv.tosjis(str)
    else
      ""
    end
  end

  def renew_skype(input="")
    case input
    when "get_head", "head", "h", "show_recent", "sr", "show"
      @skype = Skype.new(@char_code) unless @skype
    else
      @skype = Skype.new(@char_code) unless @skype
    end
  end

  def renew_chat(input, result)
    if(/^set(_chat)? .+$/ =~ input)
      @chat = result
    end
  end
end
#SJIS用
#SkypeShell.new("s").start
#UTF-8用
SkypeShell.new.start


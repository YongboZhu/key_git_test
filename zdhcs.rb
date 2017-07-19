# encoding: utf-8
require 'find'
require 'rubygems'
require 'json'
require 'rexml/document'
require 'win32ole'
require 'watir'
require 'find'

#访问一次网站并记录访问时间
def access_page(web_name,clear_cache="是")
      control = WIN32OLE.new("HttpWatch.Controller")
      web_name_ie = Watir::Browser.new
      plugin = control.IE.Attach(web_name_ie.ie)
      plugin.Log.EnableFilter(false)
      begin
        Timeout.timeout(300) do
          plugin.ClearCache() if clear_cache == "是" # 清除缓存
          plugin.Clear()
          plugin.Record()
          web_name_ie.goto(web_name)
          web_name_ie.wait
        end
      rescue
        puts "浏览器窗口未打开"
      rescue Timeout::Error
        retry
      ensure
        plugin.Stop()
      end
      ip = plugin.log.Entries.Item(0).ServerIP
      puts "本次访问网页的IP是#{ip.chomp}"
      code = plugin.Log.Entries.Item(0).StatusCode
      #puts code
      if code.to_s == "200"
         time_elapsed = plugin.Log.Entries.Summary.Time
         puts  "本次访问网页花费的时间 = #{time_elapsed}"
      else
        puts "本次访问网页失败"
        @fail_times += 1
      end
      close_browser(web_name_ie)
      @ZX_times += 1
      return time_elapsed
 end
#循环访问网站并统计结果
def access_times(web_name,times)
  @fail_times = 0
  @ZX_times = 0
  time_array = []
  times.to_i.times do
    time_array << access_page(web_name)
    time_eclipsed = calculate_average( time_array )
    access_times = @ZX_times.to_i - @fail_times.to_i
    puts "#################################"
    puts "#访问的网站是#{web_name.chomp}"
    puts "#总共尝试访问#{@ZX_times}次,访问失败#{@fail_times}次    #"
    puts "#成功访问网页#{access_times}次的平均值是#{time_eclipsed.round(3)} #"
    puts "#################################"
    sleep 2
  end
end
#关闭IE浏览器
def close_browser(web_page_ie)
      web_page_ie.close
end
#算出成功访问网站延时的平均值
def calculate_average(param_array)
      return 0 if param_array.nil?
      return 0 if param_array.size <= 0
      return param_array[0].to_f if param_array.size == 1
      min = param_array[0].to_f
      max = param_array[0].to_f
      sum_num = param_array.length - 1
      for i in 1..sum_num
        if param_array[i].to_f >= max
          max = param_array[i].to_f
        elsif param_array[i].to_f <= min
          min = param_array[i].to_f
        end
      end
      sum = 0
      for i in 0..sum_num
        numuber = param_array[i].to_f
        sum += numuber
      end
      if param_array.size > 2
        sum = sum - max - min
        sum_num = param_array.length - 2
      else
        sum_num = param_array.length
      end
      avg = sum.to_f / sum_num.to_f
      #puts "#{param_array.join(',')}的平均值是: #{avg}"
      return avg
end

puts "#################################"
puts "#            wellcome           #"
puts "#################################"
puts "请输入执行次数"
times =gets
File.open("C:\\name.txt") do |io|
 io.each_line do |line|
   access_times(line,times)
 end
end

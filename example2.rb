# -*- coding: utf-8 -*-
require 'poilite.rb'

filename = ARGV[0]

def create_testblock actions, defines
  testblock = actions.map do |action|
    testcases = defines.find_all{|d| action =~  Regexp.new('^' + d[0] + '$')}[0][1]
    target = action.scan(/「.*」/).
                  map{|x| defines.find_all{|d| x == d[0]}[0] }.
                  map{|x| x[1]} + action.scan(/"(.*?)"/).
                  flatten
    target.reduce(testcases){|testcases, define| testcases.sub('%s', define) } 
  end

  testblock.map{|block| block.split(/\n/)}.
            map{|block| block.map{|testcase| testcase.split(',',3)}}
end

def concat cells
  cells.reduce([]) do |r, x|
    if x[0] == ""
      r.last[1] += "\n" + x[1]
      r
    else
      r << x
    end
  end
end

POILite::Excel::open(filename) do |book|
  senarios = concat book.sheets[0].used_range[1..-1].map{|xs| [xs[0], xs[1]]}
  defines = concat book.sheets[1].used_range[1..-1].map{|xs| [xs[0], xs[1]]}

p  senarios.map{|senario| create_testblock senario[1].split(/\n/), defines }  
end

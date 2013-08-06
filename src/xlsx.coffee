###
#
###

# Created with Project 
# User  YuanXiangDong
# Date  13-8-6
# To change this template use File | Settings | File Templates.

logger = require 'logger'
#xlsxs = require('xlsx-writer')
nodeExcel = require('excel-export')
fs = require('fs')

XLSX_2_JSON_OPTION = {
  path:"",
  name:{
    "scenes":"_scenes.json"
  },
  fields:{
    "id":"id",
    "wuid":"wuid",
    "场景名字":"title",
    "高：（格子）":"height",
    "宽":"width"
  }
}

XLSX_OPTION = {
  path: "",
  name: "scenes"
  columns:[
    {caption:'id', type:'number'},
    {caption:'wuid', type:'string'},
    {caption:'场景名字', type:'string'},
    {caption:'高', type:'number'},
    {caption:'宽', type:'number'},
  ]
}
class Xlsx
#  json2xlsx = (option) ->

  createXlsx : (option) ->
    conf = {}
    conf.cols = option.columns
    conf.rows = []
    result = nodeExcel.execute conf
    fs.writeFileSync("./#{option.name}.xlsx", result, 'binary')

#  xlsx2json = (option)->

module.exports = Xlsx

do ->
  xlsx = new Xlsx()
  xlsx.createXlsx(XLSX_OPTION,()->

  )
#  console.log ".........."
#  conf = {}
#  conf['cols'] = [
#    {caption:'string', type:'string'},
#    {caption:'date', type:'date'},
#    {caption:'bool', type:'bool'},
#    {caption:'number 2', type:'number'}
#  ]
#  conf.rows = [
#    ['pi', (new Date(Date.UTC(2013, 4, 1))).oaDate(), true, 3.14],
#    ["e", (new Date(2012, 4, 1)).oaDate(), false, 2.7182],
#    ["M&M<>'", (new Date(Date.UTC(2013, 6, 9))).oaDate(), false, 1.2]
#  ]
#  console.log "bbbbbbbbbbbbbbbbbbbbbbbbbb"
#  console.dir nodeExcel
#  result = nodeExcel.execute(conf)
#
#  console.log "bbbbbbbbbbbbbbbbbbbbbbbbbb"
#  fs.writeFileSync('./d.xlsx', result, 'binary')
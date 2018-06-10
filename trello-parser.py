#!/usr/bin/env python3
from __future__ import print_function
import json
import os
import httplib2
import argparse
import gspread
import urllib

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from oauth2client.service_account import ServiceAccountCredentials

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'

trello_json = urllib.request.urlopen('https://trello.com/b/WGPoiaUb.json')
data = json.load(trello_json)

print("Found {} cards in {} lists.".format(len(data['cards']), len(data['lists'])))
print("Parsing...")

lists = {l['id']: l['name'] for l in data['lists']}
users = {u['id']: u['fullName'] for u in data['members']}
labels = {l['id']: l['name'] for l in data['labels']}
checklists = {l['id']: l['name'] for l in data['checklists']}

def get_card_completedWork():
    card_completedWork = {}
    for cLists in data['checklists']:
        if cLists['name'] == '本周已完成之工作內容':
            done = []
            for item in cLists['checkItems']:
                if item['state'] == 'complete':
                    done.append(item['name'])
            card_completedWork.update( {cLists['idCard']: ", \n".join(done)} )
    return card_completedWork

# 處理出席日
def get_card_attendance():
    card_attendance = {}
    for cLists in data['checklists']:
        if cLists['name'] == '本周出席實驗室頻率? (單日累計超過2小時以上，才可勾選)':
            done = []
            for item in cLists['checkItems']:
                if item['state'] == 'complete':
                    done.append(item['name'])
            card_attendance.update( {cLists['idCard']: ", \n".join(done)} )

    return card_attendance

#pluginData名稱
def get_plugin_name():
    plugin_name = {}
    plugin_fields = data['pluginData'][0]['value'].split('[')[1].split(']')[0].split('},{')
    plugin_fields[0] = plugin_fields[0][1:len(plugin_fields[0])]
    plugin_fields[len(plugin_fields)-1] = plugin_fields[len(plugin_fields)-1][0:len(plugin_fields[len(plugin_fields)-1])-1]
    for i in plugin_fields:
        name = i.split(',')[0].split(':')[1].strip('"')
        pluginId = i.split(',')[2].split(':')[1].strip('"')
        if name == '工作期間':
            pluginId = i.split(',')[3].split(':')[1].strip('"')
        plugin_name.update({pluginId:name})

    return plugin_name

def GMattendance(cardId):
    for c in data['cards'] :
        if cardId == c['id'] and len(c['pluginData']) != 0:
            pluginData = c['pluginData'][0]['value'].split('{')[2].strip('}').split(',')
            for pData in pluginData:
                dataName = pData.split(':')[0].strip('\"')
                if plugin_name[dataName] == '參加GM' and pData.split(':')[1] == 'true':
                    return "是"

def GMpresentation(cardId):
    for c in data['cards'] :
        if cardId == c['id'] and len(c['pluginData']) != 0:
            pluginData = c['pluginData'][0]['value'].split('{')[2].strip('}').split(',')
            for pData in pluginData:
                dataName = pData.split(':')[0].strip('\"')
                if plugin_name[dataName] == '上台報告' and pData.split(':')[1] == 'true':
                    return "是"

#工作起始日
def startDate(cardId):
    for c in data['cards'] :
        if cardId == c['id'] and len(c['pluginData']) != 0:
            pluginData = c['pluginData'][0]['value'].split('{')[2].strip('}').split(',')
            for pData in pluginData:
                dataName = pData.split(':')[0].strip('\"')
                if plugin_name[dataName] == '工作期間':
                    date = pData.split(':')[1][1:11]
                    return date

def get_assignor(cardId):
    for c in data['cards'] :
        if cardId == c['id'] and len(c['pluginData']) != 0:
            pluginData = c['pluginData'][0]['value'].split('{')[2].strip('}').split(',')
            for pData in pluginData:
                dataName = pData.split(':')[0].strip('\"')
                if plugin_name[dataName] == '指派者':
                    assignor = pData.split(':')[1].strip('\"')
                    return assignor

plugin_name = get_plugin_name()
card_attendance = get_card_attendance()
card_completedWork = get_card_completedWork()

parsed_cards = [{
    "name": c['name'],
    "list": lists[c['idList']],
    "description": c['desc'],
    "members": ", \n".join([u for k, u in users.items() if k in c['idMembers']]),
    "labels": ", ".join([l for k, l in labels.items() if k in c['idLabels']]),
    "completedWork": ", ".join([work for i, work in card_completedWork.items() if i == c['id']]),
    "attendance": ", ".join([day for i, day in card_attendance.items() if i == c['id']]),
    "GMattendance": GMattendance(c['id']),
    "GMpresentation": GMpresentation(c['id']),
    "startDate": startDate(c['id']),
    "assignor": get_assignor(c['id']),
    "due": c['due'][0:10] if c['due'] != None else ''

} for c in data['cards'] if c['closed'] == False ]

####################################################################################################

credentials = ServiceAccountCredentials.from_json_keyfile_dict({
  "type": "service_account",
  "project_id": "",
  "private_key_id": "7",
  "private_key": "",
  "client_email": "",
  "client_id": "",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://accounts.google.com/o/oauth2/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": ""
}
, SCOPES)
client = gspread.authorize(credentials)
http = credentials.authorize(httplib2.Http())
discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                'version=v4')
service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

flush_line_num = 50
flush_data = [[""]*10]*flush_line_num

#可直接輸出的卡片
progress=[]
assign=[]
while len(parsed_cards) != 0:
    c = parsed_cards.pop()
    if c['labels'] == '工作進度' and c['list'] != '文件':
        content = [ c['list'], c['labels'], c['members'], c['startDate'], c['description'],
                c['completedWork'], c['attendance'] , c['GMattendance'], c['GMpresentation'], c['assignor']]
        if c['list'] == 'Doing - 本週工作':
            progress.insert( 0, content)
        else:
            progress.append(content)
    elif c['labels'] == '工作指派' and c['list'] != '文件':
        content = [ c['list'], c['members'], c['due'], c['name'], c['description'], '', c['assignor'] ]
        if c['list'] == 'To Do - 指派工作':
            assign.insert( 0, content)
        else:
            assign.append(content)

spreadsheetId = '1i56ARpuOWwZKQ_utALnib8MrX0KbYaO_0B3JsKy7Di8'
rangeName = '工作進度!A2:Z'
flush = {
    "majorDimension": "DIMENSION_UNSPECIFIED",
    #"range": "A2:Z",
    "values": flush_data
    }
progress_body = {
    "majorDimension": "DIMENSION_UNSPECIFIED",
    "range": "工作進度!A2:Z",
    "values": progress
    }
assign_body = {
    "majorDimension": "DIMENSION_UNSPECIFIED",
    "range": "工作指派!A2:Z",
    "values": assign
    }
value_input_option = 'RAW'

request = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range="工作進度!A2:Z", valueInputOption=value_input_option, body=flush)
response = request.execute()
request = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range="工作指派!A2:Z", valueInputOption=value_input_option, body=flush)
response = request.execute()

request = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range=rangeName, valueInputOption=value_input_option, body=progress_body)
response = request.execute()
request = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range="工作指派!A2:Z", valueInputOption=value_input_option,     body=assign_body)
response = request.execute()

print(response)

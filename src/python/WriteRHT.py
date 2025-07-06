#!/usr/bin/python

import requests


URL = 'https://api.bdbsystem.com.br/auth/v2/user/auth'

Headers = {
    'content-type': 'application/json',
}

Payload = {
 "email": "dboliveira@brassbrasil.com.br",
 "password": "%9dE;VpvmS"
}

r = requests.post(URL, data = Payload, headers = Headers)

r.text














URL = "https://api.bdbsystem.com.br/connect/v1/worktime-record/add-activity-record"

Bearer = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9zaWQiOiIwMzExZGUwZS0yMzVkLTQzOTktOWZiNC04ZWY3ODNiODlhOTgiLCJwcmltYXJ5c2lkIjoiMWY1ZDRhOGYtODRmNC00MjdlLWE5NzAtNGI1NzhiN2Y0YTE3Iiwicm9sZSI6IltdIiwibmJmIjoxNzIwNjE4OTU5LCJleHAiOjE3MjA3MDUzNTksImlhdCI6MTcyMDYxODk1OX0.mptrclAQsJCUG2QT1Ruz7XeMSTv20qH-ejpi9eOQoeA"

Payload = {
	"projectId":"a016f2d3-7f67-4add-ac6b-fb557e7b4bc7",
	"userId":"0311de0e-235d-4399-9fb4-8ef783b89a98",
	"activityId":"792f5650-618c-11ee-a855-63d41375f55b",
	"systemId":"33676fb9-64f4-4085-b9a9-0837d05da7bb",
	"projectStageId":"64bd9557-6c1d-48c1-9752-dc878327e36e",
	"recordDate":"2024-07-10",
	"observation":"Reuinão com Vinicius Almeida (Coordenação) e Rafael Andrade (Hidraulica) sobre seleção \n",
	"workTime":"00:50",
	"externalTaskId":"",
	"isOvertime":"false",
	"activityLocation":"",
}
Headers = {
    'authorization': 'Bearer '+Bearer,    
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
    'content-type': 'application/json',
    'origin': 'https://connect.bdbsystem.com.br',
    'priority': 'u=1, i',
    'referer': 'https://connect.bdbsystem.com.br/',
    'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'x-account': '1f5d4a8f-84f4-427e-a970-4b578b7f4a17',
    'x-enviroment': 'prod',
    'x-language': 'pt-BR',
    'x-platform': 'Chrome 126.0.0.0 on Windows 10 64-bit',
    'x-tracker': '873c1c70-7558-4ef7-ba56-ba114dc21e04',
}
r = requests.post(URL, data = Payload, headers = Headers)
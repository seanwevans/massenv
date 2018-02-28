#!/usr/bin/env python

NAME = 0
STREET = 1
CITY = 2
COUNTRY = 3

first_names = []
last_names = []
street_names = []
cities = []
countries = []

with open('firstnames.txt','r') as fn:
	for line in fn:
		ls = line.split(' ')
		parsed = ls[0][0].upper() + ls[0][1:].lower()
		first_names.append(parsed)

with open('lastnames.txt','r') as ln:
	for line in ln:
		ls = line.split(' ')
		parsed = ls[0][0].upper() + ls[0][1:].lower()
		last_names.append(parsed)
		
with open('streetnames.csv','r') as sn:
	for line in sn:
		ls = line.split(',')				
		if ls[2] != '':
			parsed = ls[1][0].upper() + ls[1][1:].lower() + " " + ls[2][0].upper() + ls[2][1:].lower()
		else:
			pased = ls[1][0].upper() + ls[1][1:].lower()		
		street_names.append(parsed)

with open('cityandstate.csv','r') as cs:
	for line in cs:
		cities.append(line.split('|')[0] + " " + line.split('|')[1])
		
with open('countries.txt','r') as co:
	countries = co.read().splitlines()

if __name__ == "__main__":
	
	import xlwt
	import random
	from sys import argv

	book = xlwt.Workbook()
	sheet1 = book.add_sheet("Random Wedding Guest List")

	for guest in range(int(argv[1])):
		row = sheet1.row(guest)
		for col in range(4):
			if col == NAME:
				value = random.choice(first_names) + " " + random.choice(last_names)
			if col == STREET:
				value = str(random.randint(1,10000)) + " " + random.choice(street_names)
			if col == CITY:
				value = random.choice(cities) + " " + str(random.randint(10000,100000))
			if col == COUNTRY:
				value = random.choice(countries)
			
			row.write(col, value)

	book.save(argv[2])
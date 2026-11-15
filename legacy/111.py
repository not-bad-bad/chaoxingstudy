import requests
from lxml import etree
from docx import Document
import re
from docx.oxml.ns import qn

docum=Document()
docum.styles['Normal'].font.name = u'宋体'
docum.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

#TWFyeGlzbTI9WwogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTQ4JmNsYXNzSWQ9NzE1MyZleGFtSWQ9NjgyMDc2JmV4YW1BbnN3ZXJJZD0xNjAzNzE4JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NDgmY2xhc3NJZD03MTUzJmV4YW1JZD02ODIwNzkmZXhhbUFuc3dlcklkPTE2MDM3MjEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU0OCZjbGFzc0lkPTcxNTMmZXhhbUlkPTY4MjE1MyZleGFtQW5zd2VySWQ9MTYwMzcyMyZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTQ4JmNsYXNzSWQ9NzE1MyZleGFtSWQ9NjgyMTU5JmV4YW1BbnN3ZXJJZD0xNjAzNzk1JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NDgmY2xhc3NJZD03MTUzJmV4YW1JZD02ODIxNjQmZXhhbUFuc3dlcklkPTE2MDM3OTYmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU0OCZjbGFzc0lkPTcxNTMmZXhhbUlkPTY4MjEyOCZleGFtQW5zd2VySWQ9MTYwMzc5NyZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTQ4JmNsYXNzSWQ9NzE1MyZleGFtSWQ9NjgyMTY1JmV4YW1BbnN3ZXJJZD0xNjAzNzk4JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NDgmY2xhc3NJZD03MTUzJmV4YW1JZD02ODIwODkmZXhhbUFuc3dlcklkPTE2MDM3OTkmcHJvdG9jb2xfdj0xJwogICAgXQp4aTI9WwogICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjM0NjImY2xhc3NJZD03MTUxJmV4YW1JZD02OTQzODcmZXhhbUFuc3dlcklkPTE2MzE4NTYmcHJvdG9jb2xfdj0xJywKICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzNDYyJmNsYXNzSWQ9NzE1MSZleGFtSWQ9Njk0Mzk4JmV4YW1BbnN3ZXJJZD0xNjMxODU4JnByb3RvY29sX3Y9MScsCiAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzQ2MiZjbGFzc0lkPTcxNTEmZXhhbUlkPTY5NDM5OSZleGFtQW5zd2VySWQ9MTYzMTg1OSZwcm90b2NvbF92PTEnLAogICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjM0NjImY2xhc3NJZD03MTUxJmV4YW1JZD02OTQ0MDImZXhhbUFuc3dlcklkPTE2MzE4NjAmcHJvdG9jb2xfdj0xJywKICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzNDYyJmNsYXNzSWQ9NzE1MSZleGFtSWQ9Njk0NDA1JmV4YW1BbnN3ZXJJZD0xNjMxODYxJnByb3RvY29sX3Y9MScsCiAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzQ2MiZjbGFzc0lkPTcxNTEmZXhhbUlkPTY5NDQwNiZleGFtQW5zd2VySWQ9MTYzMTg2MiZwcm90b2NvbF92PTEnLAogICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjM0NjImY2xhc3NJZD03MTUxJmV4YW1JZD02OTQ0MTUmZXhhbUFuc3dlcklkPTE2MzE4NjMmcHJvdG9jb2xfdj0xJywKICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzNDYyJmNsYXNzSWQ9NzE1MSZleGFtSWQ9Njk0NDE5JmV4YW1BbnN3ZXJJZD0xNjMxODY0JnByb3RvY29sX3Y9MScsCiAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzQ2MiZjbGFzc0lkPTcxNTEmZXhhbUlkPTY5NDQyMSZleGFtQW5zd2VySWQ9MTYzMTg2NSZwcm90b2NvbF92PTEnLAogICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjM0NjImY2xhc3NJZD03MTUxJmV4YW1JZD02OTQ0MjMmZXhhbUFuc3dlcklkPTE2MzE4NjYmcHJvdG9jb2xfdj0xJywKICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzNDYyJmNsYXNzSWQ9NzE1MSZleGFtSWQ9Njk0NDI5JmV4YW1BbnN3ZXJJZD0xNjMxODY3JnByb3RvY29sX3Y9MScsCiAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzQ2MiZjbGFzc0lkPTcxNTEmZXhhbUlkPTY5NDQzNSZleGFtQW5zd2VySWQ9MTYzMTg2OSZwcm90b2NvbF92PTEnLAogICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjM0NjImY2xhc3NJZD03MTUxJmV4YW1JZD02OTQ0MzcmZXhhbUFuc3dlcklkPTE2MzE4NzEmcHJvdG9jb2xfdj0xJywKICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzNDYyJmNsYXNzSWQ9NzE1MSZleGFtSWQ9Njk0NDQwJmV4YW1BbnN3ZXJJZD0xNjMxODcyJnByb3RvY29sX3Y9MScKICAgIF0KbWFvMj1bCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NzMmY2xhc3NJZD03MTU1JmV4YW1JZD03NDAyNzYmZXhhbUFuc3dlcklkPTE3MzcxMTYmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU3MyZjbGFzc0lkPTcxNTUmZXhhbUlkPTc0MDI3NyZleGFtQW5zd2VySWQ9MTczNzExNyZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTczJmNsYXNzSWQ9NzE1NSZleGFtSWQ9NzQwMjk5JmV4YW1BbnN3ZXJJZD0xNzM3MTIwJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NzMmY2xhc3NJZD03MTU1JmV4YW1JZD03NDAzMDkmZXhhbUFuc3dlcklkPTE3MzcxMjEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU3MyZjbGFzc0lkPTcxNTUmZXhhbUlkPTc0MDMyNyZleGFtQW5zd2VySWQ9MTczNzEyMiZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTczJmNsYXNzSWQ9NzE1NSZleGFtSWQ9NzQwMzMxJmV4YW1BbnN3ZXJJZD0xNzM3MTIzJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NzMmY2xhc3NJZD03MTU1JmV4YW1JZD03NDAzMzYmZXhhbUFuc3dlcklkPTE3MzcxMjQmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU3MyZjbGFzc0lkPTcxNTUmZXhhbUlkPTc0MDM0MyZleGFtQW5zd2VySWQ9MTczNzEyNiZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTczJmNsYXNzSWQ9NzE1NSZleGFtSWQ9NzQwMzQ1JmV4YW1BbnN3ZXJJZD0xNzM3MTI4JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NzMmY2xhc3NJZD03MTU1JmV4YW1JZD03NDAzNTEmZXhhbUFuc3dlcklkPTE3MzcxMzEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU3MyZjbGFzc0lkPTcxNTUmZXhhbUlkPTc0MDM1NiZleGFtQW5zd2VySWQ9MTczNzEzNCZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTczJmNsYXNzSWQ9NzE1NSZleGFtSWQ9NzQwMzY3JmV4YW1BbnN3ZXJJZD0xNzM3MTM2JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTY1NzMmY2xhc3NJZD03MTU1JmV4YW1JZD03NDAzNzAmZXhhbUFuc3dlcklkPTE3MzcxMzgmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNjU3MyZjbGFzc0lkPTcxNTUmZXhhbUlkPTc0MDM3MiZleGFtQW5zd2VySWQ9MTczNzEzOSZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE2NTczJmNsYXNzSWQ9NzE1NSZleGFtSWQ9NzQwMzcxJmV4YW1BbnN3ZXJJZD0xNzM3MTQwJnByb3RvY29sX3Y9MScKICAgIF0KTWFyeGlzbTE9WwogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MjYxJmNsYXNzSWQ9NzE1NyZleGFtSWQ9Njk0NjYxJmV4YW1BbnN3ZXJJZD0xNjcxNTIzJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTcyNjEmY2xhc3NJZD03MTU3JmV4YW1JZD02OTQ3MDYmZXhhbUFuc3dlcklkPTE2NzE1NjMmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzI2MSZjbGFzc0lkPTcxNTcmZXhhbUlkPTY5NDkwMyZleGFtQW5zd2VySWQ9MTY3MTU2NCZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MjYxJmNsYXNzSWQ9NzE1NyZleGFtSWQ9Njk0OTc1JmV4YW1BbnN3ZXJJZD0xNjcxNTY1JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTcyNjEmY2xhc3NJZD03MTU3JmV4YW1JZD02OTUwMDMmZXhhbUFuc3dlcklkPTE2NzE1NzAmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzI2MSZjbGFzc0lkPTcxNTcmZXhhbUlkPTY5NDk2MyZleGFtQW5zd2VySWQ9MTY3MTU3MSZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MjYxJmNsYXNzSWQ9NzE1NyZleGFtSWQ9Njk1MDA3JmV4YW1BbnN3ZXJJZD0xNjcxNTcyJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTcyNjEmY2xhc3NJZD03MTU3JmV4YW1JZD02OTUwMDgmZXhhbUFuc3dlcklkPTE2NzE1NzMmcHJvdG9jb2xfdj0xJwogICAgXQp4aTE9WwogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzMzEwJmNsYXNzSWQ9NjY0NCZleGFtSWQ9NjAxNDA2JmV4YW1BbnN3ZXJJZD0xNDM4ODEzJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjMzMTAmY2xhc3NJZD02NjQ0JmV4YW1JZD02MDE0MDgmZXhhbUFuc3dlcklkPTE0Mzg4NTcmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzMxMCZjbGFzc0lkPTY2NDQmZXhhbUlkPTYwMTQ0MyZleGFtQW5zd2VySWQ9MTQzODg2MSZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzMzEwJmNsYXNzSWQ9NjY0NCZleGFtSWQ9NjAxNDUyJmV4YW1BbnN3ZXJJZD0xNDM4ODc4JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjMzMTAmY2xhc3NJZD02NjQ0JmV4YW1JZD02MDE0NTMmZXhhbUFuc3dlcklkPTE0Mzg4ODEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzMxMCZjbGFzc0lkPTY2NDQmZXhhbUlkPTYwMTQ1NCZleGFtQW5zd2VySWQ9MTQzODg4MiZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzMzEwJmNsYXNzSWQ9NjY0NCZleGFtSWQ9NjAxNDU1JmV4YW1BbnN3ZXJJZD0xNDM4ODgzJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjMzMTAmY2xhc3NJZD02NjQ0JmV4YW1JZD02MDE0NTYmZXhhbUFuc3dlcklkPTE0Mzg4ODUmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzMxMCZjbGFzc0lkPTY2NDQmZXhhbUlkPTYwMTQ1NyZleGFtQW5zd2VySWQ9MTQzODg5MCZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzMzEwJmNsYXNzSWQ9NjY0NCZleGFtSWQ9NjAxNDU4JmV4YW1BbnN3ZXJJZD0xNDM4ODkxJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjMzMTAmY2xhc3NJZD02NjQ0JmV4YW1JZD02MDE0NTkmZXhhbUFuc3dlcklkPTE0Mzg4OTImcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTcyMzMxMCZjbGFzc0lkPTY2NDQmZXhhbUlkPTYwMTQ2MCZleGFtQW5zd2VySWQ9MTQzODg5NCZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5NzIzMzEwJmNsYXNzSWQ9NjY0NCZleGFtSWQ9NjAxNDIwJmV4YW1BbnN3ZXJJZD0xNDM4ODk1JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjk3MjMzMTAmY2xhc3NJZD02NjQ0JmV4YW1JZD02MDE0NjEmZXhhbUFuc3dlcklkPTE0Mzg4OTcmcHJvdG9jb2xfdj0xJwogICAgIF0KbWFvMT1bCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTczMDcmY2xhc3NJZD03MDQyJmV4YW1JZD02NzA4MTgmZXhhbUFuc3dlcklkPTE1OTEwMjEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzMwNyZjbGFzc0lkPTcwNDImZXhhbUlkPTY3MDgxNyZleGFtQW5zd2VySWQ9MTU5MTAyMyZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MzA3JmNsYXNzSWQ9NzA0MiZleGFtSWQ9NjcwNzg1JmV4YW1BbnN3ZXJJZD0xNTkxMDI3JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTczMDcmY2xhc3NJZD03MDQyJmV4YW1JZD02NzExMjYmZXhhbUFuc3dlcklkPTE1OTEwMjgmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzMwNyZjbGFzc0lkPTcwNDImZXhhbUlkPTY3MTEzMCZleGFtQW5zd2VySWQ9MTU5MTAyOSZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MzA3JmNsYXNzSWQ9NzA0MiZleGFtSWQ9NjcxMTQ2JmV4YW1BbnN3ZXJJZD0xNTkxMDMxJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTczMDcmY2xhc3NJZD03MDQyJmV4YW1JZD02NzExNTUmZXhhbUFuc3dlcklkPTE1OTEwMzImcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzMwNyZjbGFzc0lkPTcwNDImZXhhbUlkPTY3MTE1NyZleGFtQW5zd2VySWQ9MTU5MTAzMyZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MzA3JmNsYXNzSWQ9NzA0MiZleGFtSWQ9NjcxMTYwJmV4YW1BbnN3ZXJJZD0xNTkxMDM2JnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTczMDcmY2xhc3NJZD03MDQyJmV4YW1JZD02NzExNjcmZXhhbUFuc3dlcklkPTE1OTEwMzcmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzMwNyZjbGFzc0lkPTcwNDImZXhhbUlkPTY3MTE3MiZleGFtQW5zd2VySWQ9MTU5MTAzOCZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MzA3JmNsYXNzSWQ9NzA0MiZleGFtSWQ9NjcxMTc1JmV4YW1BbnN3ZXJJZD0xNTkxMDQwJnByb3RvY29sX3Y9MScsCiAgICAnaHR0cHM6Ly9leG0tbWF5dWFuLWFucy5jaGFveGluZy5jb20vZXhhbS9waG9uZS9sb29rLWRldGFpbD9jb3Vyc2VJZD0yMjkzMTczMDcmY2xhc3NJZD03MDQyJmV4YW1JZD02NzExODEmZXhhbUFuc3dlcklkPTE1OTEwNDEmcHJvdG9jb2xfdj0xJywKICAgICdodHRwczovL2V4bS1tYXl1YW4tYW5zLmNoYW94aW5nLmNvbS9leGFtL3Bob25lL2xvb2stZGV0YWlsP2NvdXJzZUlkPTIyOTMxNzMwNyZjbGFzc0lkPTcwNDImZXhhbUlkPTY3MTE4NCZleGFtQW5zd2VySWQ9MTU5MTA0MiZwcm90b2NvbF92PTEnLAogICAgJ2h0dHBzOi8vZXhtLW1heXVhbi1hbnMuY2hhb3hpbmcuY29tL2V4YW0vcGhvbmUvbG9vay1kZXRhaWw/Y291cnNlSWQ9MjI5MzE3MzA3JmNsYXNzSWQ9NzA0MiZleGFtSWQ9NjcxMTc0JmV4YW1BbnN3ZXJJZD0xNTkxMDQ0JnByb3RvY29sX3Y9MScKICAgIAogICAgXQ==
#上面的是请求链接的列表，还原后放置原处即可
Marxism2=[
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682076&examAnswerId=1603718&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682079&examAnswerId=1603721&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682153&examAnswerId=1603723&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682159&examAnswerId=1603795&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682164&examAnswerId=1603796&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682128&examAnswerId=1603797&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682165&examAnswerId=1603798&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316548&classId=7153&examId=682089&examAnswerId=1603799&protocol_v=1'
    ]
xi2=[
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694387&examAnswerId=1631856&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694398&examAnswerId=1631858&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694399&examAnswerId=1631859&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694402&examAnswerId=1631860&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694405&examAnswerId=1631861&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694406&examAnswerId=1631862&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694415&examAnswerId=1631863&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694419&examAnswerId=1631864&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694421&examAnswerId=1631865&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694423&examAnswerId=1631866&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694429&examAnswerId=1631867&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694435&examAnswerId=1631869&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694437&examAnswerId=1631871&protocol_v=1',
   'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723462&classId=7151&examId=694440&examAnswerId=1631872&protocol_v=1'
    ]
mao2=[
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740276&examAnswerId=1737116&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740277&examAnswerId=1737117&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740299&examAnswerId=1737120&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740309&examAnswerId=1737121&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740327&examAnswerId=1737122&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740331&examAnswerId=1737123&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740336&examAnswerId=1737124&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740343&examAnswerId=1737126&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740345&examAnswerId=1737128&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740351&examAnswerId=1737131&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740356&examAnswerId=1737134&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740367&examAnswerId=1737136&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740370&examAnswerId=1737138&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740372&examAnswerId=1737139&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229316573&classId=7155&examId=740371&examAnswerId=1737140&protocol_v=1'
    ]
Marxism1=[
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=694661&examAnswerId=1671523&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=694706&examAnswerId=1671563&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=694903&examAnswerId=1671564&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=694975&examAnswerId=1671565&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=695003&examAnswerId=1671570&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=694963&examAnswerId=1671571&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=695007&examAnswerId=1671572&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317261&classId=7157&examId=695008&examAnswerId=1671573&protocol_v=1'
    ]
xi1=[
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601406&examAnswerId=1438813&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601408&examAnswerId=1438857&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601443&examAnswerId=1438861&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601452&examAnswerId=1438878&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601453&examAnswerId=1438881&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601454&examAnswerId=1438882&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601455&examAnswerId=1438883&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601456&examAnswerId=1438885&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601457&examAnswerId=1438890&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601458&examAnswerId=1438891&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601459&examAnswerId=1438892&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601460&examAnswerId=1438894&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601420&examAnswerId=1438895&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229723310&classId=6644&examId=601461&examAnswerId=1438897&protocol_v=1'
     ]
mao1=[
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=670818&examAnswerId=1591021&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=670817&examAnswerId=1591023&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=670785&examAnswerId=1591027&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671126&examAnswerId=1591028&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671130&examAnswerId=1591029&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671146&examAnswerId=1591031&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671155&examAnswerId=1591032&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671157&examAnswerId=1591033&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671160&examAnswerId=1591036&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671167&examAnswerId=1591037&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671172&examAnswerId=1591038&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671175&examAnswerId=1591040&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671181&examAnswerId=1591041&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671184&examAnswerId=1591042&protocol_v=1',
    'https://exm-mayuan-ans.chaoxing.com/exam/phone/look-detail?courseId=229317307&classId=7042&examId=671174&examAnswerId=1591044&protocol_v=1'

    ]
for l in Marxism1:

    head={
#此处应从浏览器上获得信息，用浏览器打开上面的任意一个链接登录后获取信息
        }

    res=requests.get(url=l,headers=head,verify=False).text
    doc=etree.HTML(res)
    num=0
    for k in range(4):
        questions=doc.xpath(f'/html/body/div[2]/h1[{k+1}]/text()')#h1加1
        if questions==[]:
            break
        pattern=re.compile('\d+')
        title_num=int(pattern.findall(doc.xpath(f'/html/body/div[2]/h1[{k+1}]/span/text()')[0][1:4])[0])
        if questions[0][-3:] in ['单选题', '多选题']:
            for _ in range(num+1,title_num+num+1):
                sequence=doc.xpath(f'/html/body/div[2]/div[{3*_-2}]/div/div/span/text()')#第二个div连续加3
                title=doc.xpath(f'/html/body/div[2]/div[{3*_-2}]/div/div/text()')[1]#第二个div连续加3
                title=title.replace('\r','').replace('\n','').replace('\t','')
                docum.add_paragraph(f'{sequence[0]}{title}')
                for i in range(4):
                    selection = f"{chr(i + 65)}、{doc.xpath(f'/html/body/div[2]/div[{3*_-2}]/div/div/ul[1]/li[{i + 1}]/label/em[2]/cc/text()')[0]}"
                    selection=selection.replace('\r','').replace('\n','').replace('\t','')
                    docum.add_paragraph(selection)
                answer=doc.xpath(f'/html/body/div[2]/div[{3*_-1}]/p[1]/text()')[0]#第二个div加3
                docum.add_paragraph(answer[2:])
        elif questions[0][-3:]=='判断题':
            for ui in range(num+1,title_num+num+1):
                sequence_adjust=doc.xpath(f'/html/body/div[2]/div[{3*ui-2}]/div/div/span/text()')[0]
                title_adjust=doc.xpath(f'/html/body/div[2]/div[{3*ui-2}]/div/div/text()')[1]
                title_adjust=title_adjust.replace('\r','').replace('\n','').replace('\t','')
                answer_adjust = f"答案：{doc.xpath(f'/html/body/div[2]/div[{3*ui-1}]/p[1]/span/text()')[0]}"
                docum.add_paragraph(f'{sequence_adjust}{title_adjust}')
                docum.add_paragraph(answer_adjust)
        else:
            for yu in range(num+1,title_num+num+1):
                sequence_write=doc.xpath(f'/html/body/div[2]/div[{3*yu-2}]/div/div/span/text()')[0]
                title_write=doc.xpath(f'/html/body/div[2]/div[{3*yu-2}]/div/div/text()')[1]
                title_write=title_write.replace('\r','').replace('\n','').replace('\t','')
                answer_write=doc.xpath(f'/html/body/div[2]/div[{3*yu-1}]/div[1]/ul/li/label/p/text()')[0]
                docum.add_paragraph(f'{sequence_write}{title_write}')
                docum.add_paragraph(f'答案：{answer_write}')
        num+=int(pattern.findall(doc.xpath(f'/html/body/div[2]/h1[{k+1}]/span/text()')[0][1:4])[0])
    print(num)
docum.save('1.docx')
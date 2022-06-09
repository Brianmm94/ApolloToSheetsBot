from distutils.core import setup

setup(name='Apollo To Sheets Bot',
      version='1.0',
      author='Brian Moran',
      description='Parses Apollo Bot messages and puts them into a Google Sheets document',
      url='https://github.com/Brianmm94/ApolloToSheetsBot',
      install_requires=['beautifulsoup4==4.11.1', 'chat_exporter==1.7.2', 'discord.py==1.7.3', 'google_api_python_client==2.50.0', 'google_auth_oauthlib==0.4.6', \
                        'protobuf==4.21.1', 'python-dotenv==0.20.0', 'python_dateutil==2.8.2']
     )
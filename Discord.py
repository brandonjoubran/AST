import discord
from discord.ext import commands
from StockScraper import StockScraper
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy

DISCORD_TOKEN = 'NzE5NjE0MjQ1Nzg3OTI2NTQ3.Xt5_JQ.alEzpaS_NMQPwnopPMTZoVEFOe8'
DISCORD_GUILD = 'Auto-Stock Trader'
client = commands.Bot(command_prefix='!')
date = datetime.today().strftime('%Y-%m-%d')
queue = []

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('ast.json', scope)
client_sp = gspread.authorize(creds)
sheet = client_sp.open('Clients').sheet1

plan_info = {"weekly": "You are currently paying $10/week and this allows for 5 tickers per day. Today you have used {} tickers, which means you can !run on {} more.",
             "monthly": "You are currently paying $30/month and this allows for unlimited tickers per day",
             "yearly": "You are currently paying $300/year and this allows for unlimited tickers per day.",
             "N/A": "You are currently not subscribed to any plan. To change this, contact an admin!"}


# Checking if user is a client
def is_client(author):

    print(sheet.row_values(1))
    print(author)
    if author not in sheet.row_values(1):
        print("AUTHOR NOT IN.")
        return False

    return True

# Helper function checking the amount of stocks that were queried

def amount_queried(author):
    for i in range(len(sheet.row_values(1))):
        if(sheet.row_values(1)[i] == author):
            return sheet.row_values(3)[i]

# Checking if a user is allowed to query the amount of stocks they queried
def can_query_amount(author, amount):

    for i in range(len(sheet.row_values(1))):
        print(sheet.row_values(5)[i])
        print(amount)
        print(sheet.row_values(1)[i])
        print(author)
        if(sheet.row_values(1)[i] == author):
            if(sheet.row_values(5)[i] == 'inf'):
                #sheet.update_cell(3, i + 1, str(int(sheet.row_values(3)[i]) + amount))
                return True
            if (int(sheet.row_values(5)[i]) < amount):
                return False
            if (int(sheet.row_values(3)[i]) + amount <= int(sheet.row_values(5)[i])):
                print("abcdefg")
                #sheet.cell(3, 1).value = str(int(sheet.row_values(3)[i]) + amount)
                sheet.update_cell(3, i+1, str(int(sheet.row_values(3)[i]) + amount))
            else:
                return False
            return True
    return False

# Returning a user's plan

def get_user_plan(author):
    for i in range(len(sheet.row_values(1))):
        if (sheet.row_values(1)[i] == author):
            return sheet.row_values(6)[i]

    return "N/A"

# Updating client database 

def update_daily_info(author):
    for i in range(len(sheet.row_values(1))):
        if(sheet.row_values(1)[i] == author):
            if (sheet.row_values(2)[i] != date):
                sheet.update_cell(2, i + 1, date)
                sheet.update_cell(3, i+1, 0)
                sheet.update_cell(4, i + 1, "")

def update_daily_tickers(author, new):
    for i in range(len(sheet.row_values(1))):
        if(sheet.row_values(1)[i] == author):
            print("RIGHT ERE")
            print(new)
            print(sheet.cell(4, i).value)
            print(sheet.cell(4, i+1).value)
            current_tickers = sheet.cell(4, i+1).value
            print(current_tickers)
            if(current_tickers != "" or current_tickers != None):
                current_tickers += ','
                current_tickers += new
                sheet.update_cell(4, i + 1, current_tickers)

def update_database(row, uid):
    database = client_sp.open('Clients').worksheet(str(uid))
    database.append_row(row)
    print(database.row_count)
    print(database.get_all_values())
    print(database.row_values(1))

@client.event
async def on_ready():

    print(f'{client.user} has connected to Discord!')
    for guild in client.guilds:
        if guild.name == DISCORD_GUILD:
            break
    print(
        f'{client.user} is connected to the following guild:\n'
        f'{guild.name}(id: {guild.id})'
    )

    members = '\n - '.join([member.name for member in guild.members])
    print(f'Guild Members:\n - {members}')

@client.command(name="greeting",  help='jus greeting')
async def greeting(ctx):
    print("here")
    response = 'I have arrived.'
    await ctx.send(response)

@client.command(name="members",  help='jus members')
async def members(ctx):
    for guild in client.guilds:
        if guild.name == DISCORD_GUILD:
            break
    response = '\n - '.join([member.name for member in guild.members])
    await ctx.send(f'Guild Members:\n - {response}')

@client.command(name="run1",  help='Dont use does nothing')
async def run1(ctx, *stocks):
    print(ctx.message.author)
    #await client.send_file(ctx.message.author, "this.xlsx", filename="Hello", content="Message test")
    print(ctx.message.author.name)
    await ctx.message.author.send(file=discord.File('this.xlsx'))

@client.command(name="test",  help="testing function for development (do not use)")
async def test(ctx):
    try:
        database = client_sp.open('Clients').worksheet(str(ctx.message.author.id))
    except:
        await ctx.send("{}: You're not a client, if you feel this is a mistake contact an admin.".format(ctx.author.mention))
        return

    tens = {'k': 10e2, 'K': 10e2, 'm': 10e5, 'M': 10e5, 'b': 10e8, 'B': 10e8}
    f = lambda x: int(float(x[:-1]) * tens[x[-1]])

    first_cat = [0.25, 0.75, 0.3, 0.1, 0.9, 0.65, 0.32]
    second_cat = [0.6, 0.5, 0.85, 0.15, 0.42, 0.14, 0.94]
    #third_cat = [0.1, 0.65]
    #fourth_cat = [0.3, 0.125]
    #fifth_cat = [0.93, 0.98]
    #sixth_cat = [0.53, 0.21]
    #seventh_cat = [0.1, 0.9]
    x = numpy.arange(len(first_cat))
    #plt.ioff()
    bar_width = 0.40
    plt.bar(x, first_cat, width=bar_width, color='green')
    plt.bar(x+bar_width, second_cat, width=bar_width, color='red')
    #plt.bar(x + bar_width*2, third_cat, width=bar_width, color='green')
    #plt.bar(x + bar_width * 3, fourth_cat, width=bar_width, color='red')
    #plt.bar(x + bar_width * 4, fifth_cat, width=bar_width, color='orange')

    plt.xticks(x + bar_width/2, ['0-50M', '50M-100M', '100M-300M', '300M-499.9M', '500M-999.9M', '1B-9.99B', '10B+'], rotation=45)
    plt.title = "Market Cap to Spike/Fail"
    plt.xlabel = 'xaxis'
    plt.ylabel = 'yaxis'
    #plt.close()
    plt.savefig('test.pdf', bbox_inches="tight")
    #plt.show()


    '''for record in database.get_all_records():
        print(record['Market Cap'])
        print(f(record['Market Cap']))
        if(f(record['Market Cap']) > 0 and f(record['Market Cap']) < 49999999):
            first_cat += [f(record['Market Cap'])]
        elif(f(record['Market Cap']) > 50000000 and f(record['Market Cap']) < 99999999):
            second_cat += [f(record['Market Cap'])]
        elif (f(record['Market Cap']) > 100000000 and f(record['Market Cap']) < 299999999):
            third_cat += [f(record['Market Cap'])]
        elif (f(record['Market Cap']) > 300000000 and f(record['Market Cap']) < 499999999):
            fourth_cat += [f(record['Market Cap'])]
        elif (f(record['Market Cap']) > 500000000 and f(record['Market Cap']) < 999999999):
            fifth_cat += [f(record['Market Cap'])]
        elif (f(record['Market Cap']) > 1000000000 and f(record['Market Cap']) < 9999999999):
            sixth_cat += [f(record['Market Cap'])]
        elif (f(record['Market Cap']) > 10000000000):
            seventh_cat += [f(record['Market Cap'])]'''

    print(first_cat)
    print(second_cat)
    #print(third_cat)
    #print(fourth_cat)
    #print(fifth_cat)
    #print(sixth_cat)
    #print(seventh_cat)
    #print(np.median(first_cat))



@client.command(name="report",  help="Get a report of ticker patterns (being developed)")
async def report(ctx):
    #update_database(['a', 'b', 'c', 'd'], ctx.message.author.id)
    #str(ctx.message.author.id)
    try:
        database = client_sp.open('Clients').worksheet('444482373308907520')
    except:
        await ctx.send("{}: You're not a client, if you feel this is a mistake contact an admin.".format(ctx.author.mention))
        return
    #print(database.get_all_records())
    records = database.get_all_records()
    await ctx.send("You're query history is being rendered...")
    dict = {}
    for record in records:
        try:
            dict[record['Stocks'].upper()] += 1
        except:
            dict[record['Stocks'].upper()] = 1

    response = ''
    for ticker, num in dict.items():
        response += "{}: {} \n".format(ticker, num)
    await ctx.send(response)

    print(dict)


@client.command(name="plan",  help="User's Current Plan Information")
async def plan(ctx):
    if(get_user_plan(ctx.message.author.name) == 'weekly'):
        await ctx.message.author.send(plan_info[get_user_plan(ctx.message.author.name)].format(amount_queried(ctx.message.author.name), 5-int(amount_queried(ctx.message.author.name))))
        return
    await ctx.message.author.send(plan_info[get_user_plan(ctx.message.author.name)].format(amount_queried(ctx.message.author.name)))

@client.command(name="run", help='Enter the tickers')
async def run(ctx, *stocks):
    update_daily_info(ctx.message.author.name)
    print(queue)
    stock_list = []
    stock_list += [word for word in stocks]
    stocks_string = ', '.join([stock for stock in stock_list])
    print(stocks_string)
    print(stocks_string.split(','))

    if(is_client(ctx.message.author.name) == False):
        await ctx.send("You are not a client. If you feel this is a mistake, contact an admin.")
        return
    if(can_query_amount(ctx.message.author.name, len(stock_list)) == False):
        await ctx.send("You are asking for more tickers than your current plan allows. Do !plan and we will DM you your current plan information.")
        return

    filename = ctx.message.author.name
    queue.append([filename, stock_list])
    # await client.send_file(ctx.message.author, "this.xlsx", filename="Hello", content="Message test")
    print(stock_list)
    await ctx.send("The stocks you entered were ({}):".format(len(stock_list)))
    for tuple in queue:
        for stock in tuple[1]:
            await ctx.send(stock)
        StockScraper(tuple[1], tuple[0], str(ctx.message.author.id))
        await ctx.message.author.send(file=discord.File(tuple[0]+".xlsx"))
        #await ctx.message.author.send(file=discord.File("Isaac18.xlsx"))
        update_daily_tickers(ctx.message.author.name, stocks_string)
        queue.pop(0)
client.run(DISCORD_TOKEN)



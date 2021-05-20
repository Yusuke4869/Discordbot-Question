import os
import discord
import glob
import xlrd

# token.txtファイルからTOKENの読み込み
with open("token.txt") as f:
    TOKEN = f.read()

client = discord.Client()

@client.event
async def on_ready():
    print("logged in\n")
    await client.change_presence(activity=discord.Game(name="please mention", type=1))

doing_list = []

@client.event
async def on_message(message):

    if message.author.bot:
        return

    if not client.user in message.mentions:
        return

    select_message = "1: 1問1答を始める\n2: help\n3: Excelテンプレートファイルを入手する"

    playing = 0
    if message.guild.id in doing_list:
        select_message = f"{select_message}\n4: 1問1答を強制終了する"
        playing = 1

    embed=discord.Embed(title="1問1答", description="メニュー", color=0xdeb887)
    embed.add_field(name="モードを選択してください", value=select_message, inline=False)
    select_send_message = await message.channel.send(embed=embed)

    await select_send_message.add_reaction("1️⃣")
    await select_send_message.add_reaction("2️⃣")
    await select_send_message.add_reaction("3️⃣")

    if playing == 1:
        await select_send_message.add_reaction("4️⃣")

    def select_check(reaction, user):
        emoji = reaction.emoji
        if user != message.author:
            pass
        else:
            return emoji == "1️⃣" or emoji == "2️⃣" or emoji =="3️⃣" or emoji == "4️⃣"

    reaction, user= await client.wait_for("reaction_add", check=select_check)

    await select_send_message.clear_reactions()

    if reaction.emoji == "1️⃣":

        if message.guild.id in doing_list:
            return await message.channel.send("すでに開始されています\n開始する場合には強制終了してください")

        start_message = await message.channel.send(f"{message.author.display_name}さん、1問1答を始めますか？")

        await start_message.add_reaction("⭕")
        await start_message.add_reaction("❌")

        def first_check(reaction, user):
            emoji = reaction.emoji
            if user != message.author:
                pass
            else:
                return emoji == "⭕" or emoji == "❌"

        reaction, user= await client.wait_for("reaction_add", check=first_check)

        await start_message.clear_reactions()

        if reaction.emoji == "⭕":
            doing_list.append(message.guild.id)
        elif reaction.emoji == "❌":
            return await message.channel.send("終了します")
        else:
            return await message.channel.send("エラー\nはじめからやり直してください")

        files_name = ""
        file_num = 0
        files_list = []
        files_name_list = []
        files = glob.glob('./files/*.xlsx')
        for file in files:
            file_name = os.path.split(file)[1].rstrip(".xlsx")
            judge_file_name = file_name.lower()
            if "template" in judge_file_name:
                pass
            else:
                file_num = file_num + 1
                files_name = f"{files_name}{file_num}: {file_name}\n"
                files_list.append(file)
                files_name_list.append(file_name)

        file_embed = discord.Embed(title="ファイル選択", description="1問1答を行うデータを選択してください", color=0x708090)
        file_embed.add_field(name="一覧", value=f"{files_name}")
        file_embed.set_footer(text="template と記載しているファイルは除外しています")
        await message.channel.send(embed=file_embed)

        def second_check(second_message):
            if second_message.author != message.author:
                pass
            else:
                if second_message.content.isdecimal():
                    if int(second_message.content) > 0 and int(second_message.content) <= file_num:
                        return second_message
                    else:
                        pass
                else:
                    pass

        second_message = await client.wait_for("message", check=second_check)
        open_file_num = int(second_message.content) - 1
        await message.channel.send(f"{files_name_list[open_file_num]} で行います")

        sheet_names_list = []
        excel_file = xlrd.open_workbook(files_list[open_file_num])
        sheet_names_list = excel_file.sheet_names()
        sheet = excel_file.sheet_by_name(sheet_names_list[0])

        mode_embed = discord.Embed(title="モード選択", description="モードを選択してください", color=0x708090)
        mode_embed.add_field(name="Aモード", value=f"{sheet.cell_value(1, 1)}")
        mode_embed.add_field(name="Bモード", value=f"{sheet.cell_value(1, 2)}")
        mode_select_message = await message.channel.send(embed=mode_embed)

        await mode_select_message.add_reaction("🅰")
        await mode_select_message.add_reaction("🅱")

        def third_check(reaction, user):
            emoji = reaction.emoji
            if user != message.author:
                pass
            else:
                return emoji == "🅰" or emoji == "🅱"

        reaction, user= await client.wait_for("reaction_add", check=third_check)

        await mode_select_message.clear_reactions()

        mode = 0
        roop = 0

        if reaction.emoji == "🅰":
            await message.channel.send("Aモードで行います")
            mode = 1
            roop = 1
        elif reaction.emoji == "🅱":
            await message.channel.send("Bモードで行います")
            mode = 2
            roop = 1
        else:
            return await message.channel.send("エラー\nはじめからやり直してください")
        
        total = 0
        correct_answer = 0
        wrong_answer = 0
        skip = 0

        question = 2

        correct_answer_color = 0x7cfc00
        wrong_answer_color = 0xff1493
        skip_color = 0xffd700
        end_color = 0x00bfff
        question_color = 0x4169e1

        def answer_check(answer_message):
            if answer_message.author != message.author:
                pass
            else:
                return answer_message

        def next_check(reaction, user):
            emoji = reaction.emoji
            if user != message.author:
                pass
            else:
                return emoji == "▶️" or emoji == "⭕"
        
        while roop == 1:

            if mode == 1:
                if sheet.cell_value(question, 1).lower() == "skip" or sheet.cell_value(question, 2).lower() == "skip":
                    pass
                elif sheet.cell_value(question, 1) == "":
                    end_embed = discord.Embed(title="終了です", description="", color=end_color)
                    end_embed.add_field(name="結果", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                    await message.channel.send(embed=end_embed)
                    roop = 0
                    if message.guild.id in doing_list:
                        doing_list.remove(message.guild.id)
                else:
                    total = total + 1

                    problem_embed = discord.Embed(title="問題", description="", color=question_color)
                    problem_embed.add_field(name=f"第{total}問目", value=sheet.cell_value(question, 1))
                    problem_embed.set_footer(text="スキップする場合は`skip`\n終了する場合は`end`と入力してください")
                    await message.channel.send(embed=problem_embed)

                    answer_message = await client.wait_for("message", check=answer_check)

                    if answer_message.content == sheet.cell_value(question, 2):
                        correct_answer = correct_answer + 1
                        answer_embed = discord.Embed(title="正解です！", description="", color=correct_answer_color)
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=answer_embed)
                    elif answer_message.content.lower() == "skip":
                        skip = skip + 1
                        answer_embed = discord.Embed(title="スキップしました！", description="", color=skip_color)
                        answer_embed.add_field(name="正答", value=sheet.cell_value(question, 2))
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=answer_embed)
                    elif answer_message.content.lower() == "end":
                        end_embed = discord.Embed(title="終了です", description="", color=end_color)
                        end_embed.add_field(name="結果", value=f"合計 {total - 1}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=end_embed)
                        roop = 0
                        if message.guild.id in doing_list:
                            doing_list.remove(message.guild.id)
                    else:
                        wrong_answer = wrong_answer + 1
                        answer_embed = discord.Embed(title="不正解です", description="次の問題に進むには▶️を\n正解の場合は⭕を押してください", color=wrong_answer_color)
                        answer_embed.add_field(name="あなたの解答", value=answer_message.content)
                        answer_embed.add_field(name="正答", value=sheet.cell_value(question, 2))
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問", inline=False)
                        wrong_answer_message = await message.channel.send(embed=answer_embed)
                        
                        await wrong_answer_message.add_reaction("▶️")
                        await wrong_answer_message.add_reaction("⭕")
                        reaction, user= await client.wait_for("reaction_add", check=next_check)

                        if reaction.emoji == "▶️":
                            pass
                        elif reaction.emoji == "⭕":
                            wrong_answer = wrong_answer - 1
                            correct_answer = correct_answer + 1

                            answer_embed = discord.Embed(title="正解です！", description="訂正しました", color=correct_answer_color)
                            answer_embed.add_field(name="正答", value=sheet.cell_value(question, 2))
                            answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                            await wrong_answer_message.edit(embed=answer_embed, content=None)

                        await wrong_answer_message.clear_reactions()

                question = question + 1

            elif mode == 2:
                if sheet.cell_value(question, 1).lower() == "skip" or sheet.cell_value(question, 2).lower() == "skip":
                    pass
                elif sheet.cell_value(question, 2) == "":
                    end_embed = discord.Embed(title="終了です", description="", color=end_color)
                    end_embed.add_field(name="結果", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                    await message.channel.send(embed=end_embed)
                    roop = 0
                    if message.guild.id in doing_list:
                        doing_list.remove(message.guild.id)
                else:
                    total = total + 1

                    problem_embed = discord.Embed(title="問題", description="", color=question_color)
                    problem_embed.add_field(name=f"第{total}問目", value=sheet.cell_value(question, 2))
                    problem_embed.set_footer(text="スキップする場合は`skip`\n終了する場合は`end`と入力してください")
                    await message.channel.send(embed=problem_embed)

                    answer_message = await client.wait_for("message", check=answer_check)

                    if answer_message.content == sheet.cell_value(question, 1):
                        correct_answer = correct_answer + 1
                        answer_embed = discord.Embed(title="正解です！", description="", color=correct_answer_color)
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=answer_embed)
                    elif answer_message.content.lower() == "skip":
                        skip = skip + 1
                        answer_embed = discord.Embed(title="スキップしました！", description="", color=skip_color)
                        answer_embed.add_field(name="正答", value=sheet.cell_value(question, 1))
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=answer_embed)
                    elif answer_message.content.lower() == "end":
                        end_embed = discord.Embed(title="終了です", description="", color=end_color)
                        end_embed.add_field(name="結果", value=f"合計 {total - 1}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                        await message.channel.send(embed=end_embed)
                        roop = 0
                        if message.guild.id in doing_list:
                            doing_list.remove(message.guild.id)
                    else:
                        wrong_answer = wrong_answer + 1
                        answer_embed = discord.Embed(title="不正解です", description="次の問題に進むには▶️を\n正解の場合は⭕を押してください", color=wrong_answer_color)
                        answer_embed.add_field(name="あなたの解答", value=answer_message.content)
                        answer_embed.add_field(name="正答", value=sheet.cell_value(question, 1))
                        answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問", inline=False)
                        wrong_answer_message = await message.channel.send(embed=answer_embed)
                        
                        await wrong_answer_message.add_reaction("▶️")
                        await wrong_answer_message.add_reaction("⭕")
                        reaction, user= await client.wait_for("reaction_add", check=next_check)

                        if reaction.emoji == "▶️":
                            pass
                        elif reaction.emoji == "⭕":
                            wrong_answer = wrong_answer - 1
                            correct_answer = correct_answer + 1

                            answer_embed = discord.Embed(title="正解です！", description="訂正しました", color=correct_answer_color)
                            answer_embed.add_field(name="正答", value=sheet.cell_value(question, 1))
                            answer_embed.add_field(name="現在の進行状況", value=f"合計 {total}問\n正解 {correct_answer}問\n誤答 {wrong_answer}問\nスキップ {skip}問")
                            await wrong_answer_message.edit(embed=answer_embed, content=None)

                        await wrong_answer_message.clear_reactions()

                question = question + 1
                
            else:
                roop = 0
                if message.guild.id in doing_list:
                    doing_list.remove(message.guild.id)
    
    elif reaction.emoji == "2️⃣":
        await message.channel.send("```\n1問1答Bot\nエクセルファイルを使った1問1答Botです\nメンションするとモードを選択できます\n```")

    elif reaction.emoji == "3️⃣":
        file_send = 0
        files = glob.glob('./files/*.xlsx')
        for file in files:
            file_name = os.path.split(file)[1].rstrip(".xlsx")
            judge_file_name = file_name.lower()
            if "template" == judge_file_name:
                await message.channel.send(file=discord.File(file))
                file_send = 1
                break
        if file_send == 0:
            await message.channel.send("ファイルが見つかりません")

    elif reaction.emoji == "4️⃣":
        close_embed=discord.Embed(title="強制終了", description="", color=0x708090)
        close_embed.add_field(name="強制終了しますか？", value="終了する場合は⭕\nしない場合は❌を選択してください", inline=False)
        close_embed.set_footer(text="他人の邪魔をすることはやめましょう")
        close_send_message = await message.channel.send(embed=close_embed)

        await close_send_message.add_reaction("⭕")
        await close_send_message.add_reaction("❌")

        def close_check(reaction, user):
            emoji = reaction.emoji
            if user != message.author:
                pass
            else:
                return emoji == "⭕" or emoji == "❌"

        reaction, user= await client.wait_for("reaction_add", check=close_check)

        await close_send_message.clear_reactions()

        if reaction.emoji == "⭕":
            if message.guild.id in doing_list:
                doing_list.remove(message.guild.id)
                await message.channel.send("強制終了しました")
            else:
                await message.channel.send("終了に失敗しました")
        elif reaction.emoji == "❌":
            await message.channel.send("強制終了が拒否されました")
        else:
            await message.channel.send("エラー")

    else:
        await message.channel.send("エラー")

if __name__ == "__main__":
    client.run(TOKEN)
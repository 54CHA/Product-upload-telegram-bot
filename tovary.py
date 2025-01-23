import asyncio
import openpyxl
import io
import os
import aiohttp
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
import nest_asyncio

nest_asyncio.apply()

load_dotenv()

TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')
STRAPI_API_TOKEN = os.getenv('STRAPI_API_TOKEN')
STRAPI_API_URL = os.getenv('STRAPI_API_URL')

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [
            InlineKeyboardButton("📥 Загрузить товары", callback_data='upload_products'),
            InlineKeyboardButton("📤 Выгрузить товары", callback_data='export_products')
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        'Добро пожаловать! Выберите действие:',
        reply_markup=reply_markup
    )

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == 'upload_products':
        await query.edit_message_text(
            text="Пожалуйста, отправьте Excel файл с товарами.\n\n"
                 "Excel файл должен содержать следующие столбцы:\n"
                 "A: Название\n"
                 "B: Артикул\n"
                 "C: Описание\n"
                 "D: ID категории\n"
                 "E: ID подкатегории\n"
                 "F: ID бренда\n"
                 "G: Спецификации (краткие)\n"
                 "H: Спецификации (подробные)\n"
                 "I: Ссылка где купить"
        )
    elif query.data == 'export_products':
        await export_products(update, context)

async def process_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = await context.bot.get_file(update.message.document.file_id)
        file_bytes = await file.download_as_bytearray()
        
        products = extract_data_from_excel(file_bytes)
        
        if not products:
            await update.message.reply_text("В Excel файле не найдено товаров или произошла ошибка при обработке.")
            return

        await update.message.reply_text(f"Найдено {len(products)} товаров. Начинаю загрузку в Strapi...")
        
        async with aiohttp.ClientSession() as session:
            success_count = 0
            duplicate_count = 0
            error_count = 0
            
            for product in products:
                result = await create_product_in_strapi(session, product, None)
                if result['success']:
                    success_count += 1
                    await update.message.reply_text(f"✅ Создан: {product['name']}")
                else:
                    if result['reason'] == 'duplicate':
                        duplicate_count += 1
                        await update.message.reply_text(f"⚠️ Пропущен дубликат: {product['name']}")
                    else:
                        error_count += 1
                        await update.message.reply_text(f"❌ Ошибка создания: {product['name']}")
        
        await update.message.reply_text(
            f"Загрузка завершена!\n"
            f"✅ Успешно создано: {success_count} товаров\n"
            f"⚠️ Пропущено дубликатов: {duplicate_count} товаров\n"
            f"❌ Ошибок: {error_count} товаров"
        )
    
    except Exception as e:
        await update.message.reply_text(f"Произошла ошибка: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        'Пожалуйста, отправьте Excel файл с данными о товарах.'
    )

async def create_product_in_strapi(session, product_data, image_id):
    try:
        headers = {
            'Authorization': f'Bearer {STRAPI_API_TOKEN}',
            'Content-Type': 'application/json'
        }

        # URL encode the article number for the query
        encoded_article = product_data["article"].replace(' ', '%20')
        
        # Check if product with this article number already exists
        async with session.get(
            f'{STRAPI_API_URL}/api/catalog-products?filters[articleNumber][$eq]={encoded_article}',
            headers=headers
        ) as response:
            if response.status == 200:
                existing_products = await response.json()
                if existing_products.get('data') and len(existing_products['data']) > 0:
                    return {'success': False, 'reason': 'duplicate'}
            else:
                print(f"Error checking for duplicates: {await response.text()}")

        slug = product_data['name'].lower().replace(' ', '-')
        
        # Convert specifications to the correct format
        specs = product_data.get('specifications', [])
        detailed_specs = product_data.get('detailedSpecifications', [])

        data = {
            "data": {
                "name": product_data['name'],
                "slug": slug,
                "articleNumber": product_data['article'],
                "description": product_data['description'],
                "specifications": specs,
                "detailedSpecifications": detailed_specs,
                "whereToBuyLink": product_data.get('whereToBuyLink', ""),
                "publishedAt": None
            }
        }

        # Add relations
        if product_data.get('category'):
            data["data"]["category"] = product_data['category']
        if product_data.get('subcategory'):
            data["data"]["subcategory"] = product_data['subcategory']
        if product_data.get('brand'):
            data["data"]["brand"] = product_data['brand']

        # Add image if provided and not None
        if image_id:
            data["data"]["images"] = [image_id]

        print(f"Sending data to Strapi: {data}")

        async with session.post(
            f'{STRAPI_API_URL}/api/catalog-products',
            json=data,
            headers=headers
        ) as response:
            response_text = await response.text()
            print(f"Response from Strapi: {response_text}")
            if response.status not in [200, 201]:
                return {'success': False, 'reason': 'api_error'}
            return {'success': True}
    except Exception as e:
        print(f"Error creating product: {e}")
        return {'success': False, 'reason': 'exception', 'error': str(e)}

def extract_data_from_excel(excel_bytes):
    temp_file = 'temp_excel.xlsx'
    with open(temp_file, 'wb') as f:
        f.write(excel_bytes)

    try:
        workbook = openpyxl.load_workbook(temp_file)
        sheet = workbook.active

        products_data = []
        for row_idx in range(2, sheet.max_row + 1):
            if not sheet.cell(row=row_idx, column=1).value:
                continue

            # Process specifications (column G)
            spec_value = str(sheet.cell(row=row_idx, column=7).value or '')
            specs = []
            if spec_value:
                spec_parts = spec_value.split(',')
                for i, part in enumerate(spec_parts):
                    if ':' in part:
                        name, value = part.split(':', 1)
                        specs.append({
                            "label": name.strip(),
                            "value": value.strip()
                        })
                    else:
                        specs.append({
                            "label": f"Specification {i+1}",
                            "value": part.strip()
                        })

            # Process detailed specifications (column H)
            detailed_spec_value = str(sheet.cell(row=row_idx, column=8).value or '')
            detailed_specs = []
            if detailed_spec_value:
                detailed_spec_parts = detailed_spec_value.split(',')
                for i, part in enumerate(detailed_spec_parts):
                    if ':' in part:
                        name, value = part.split(':', 1)
                        detailed_specs.append({
                            "label": name.strip(),
                            "value": value.strip()
                        })
                    else:
                        detailed_specs.append({
                            "label": f"Specification {i+1}",
                            "value": part.strip()
                        })

            # Ensure at least one specification exists for both
            if not specs:
                specs = [{
                    "label": "General",
                    "value": "Not specified"
                }]
            if not detailed_specs:
                detailed_specs = [{
                    "label": "General",
                    "value": "Not specified"
                }]

            product = {
                'name': str(sheet.cell(row=row_idx, column=1).value or '').strip(),
                'article': str(sheet.cell(row=row_idx, column=2).value or '').strip(),
                'description': str(sheet.cell(row=row_idx, column=3).value or '').strip(),
                'category': int(sheet.cell(row=row_idx, column=4).value or 0),
                'subcategory': int(sheet.cell(row=row_idx, column=5).value or 0),
                'brand': int(sheet.cell(row=row_idx, column=6).value or 0),
                'specifications': specs,
                'detailedSpecifications': detailed_specs,
                'whereToBuyLink': str(sheet.cell(row=row_idx, column=9).value or '').strip()
            }
            
            if product['name'] and product['article'] and product['category'] and product['whereToBuyLink']:
                products_data.append(product)

        return products_data

    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return []
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)

async def export_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.callback_query.edit_message_text("⏳ Загружаю данные из Strapi...")
        
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        
        headers = ["Название", "Артикул", "Описание", "ID категории", "ID подкатегории", 
                  "ID бренда", "Спецификации (краткие)", "Спецификации (подробные)", "Ссылка где купить"]
        for col, header in enumerate(headers, 1):
            sheet.cell(row=1, column=col, value=header)
        
        async with aiohttp.ClientSession() as session:
            headers = {
                'Authorization': f'Bearer {STRAPI_API_TOKEN}',
                'Content-Type': 'application/json'
            }
            
            async with session.get(
                f'{STRAPI_API_URL}/api/catalog-products?populate=*',
                headers=headers
            ) as response:
                if response.status != 200:
                    await update.callback_query.edit_message_text("❌ Ошибка при получении данных из Strapi")
                    return
                
                data = await response.json()
                products = data.get('data', [])
                
                for row, product in enumerate(products, 2):
                    attrs = product.get('attributes', {})
                    
                    # Format specifications
                    specs = attrs.get('specifications', [])
                    specs_str = ', '.join([f"{spec.get('label', '')}:{spec.get('value', '')}" 
                                         for spec in specs])
                    
                    # Format detailed specifications
                    detailed_specs = attrs.get('detailedSpecifications', [])
                    detailed_specs_str = ', '.join([f"{spec.get('label', '')}:{spec.get('value', '')}" 
                                                  for spec in detailed_specs])

                    sheet.cell(row=row, column=1, value=attrs.get('name', ''))
                    sheet.cell(row=row, column=2, value=attrs.get('articleNumber', ''))
                    sheet.cell(row=row, column=3, value=attrs.get('description', ''))
                    sheet.cell(row=row, column=4, value=attrs.get('category', {}).get('data', {}).get('id', ''))
                    sheet.cell(row=row, column=5, value=attrs.get('subcategory', {}).get('data', {}).get('id', ''))
                    sheet.cell(row=row, column=6, value=attrs.get('brand', {}).get('data', {}).get('id', ''))
                    sheet.cell(row=row, column=7, value=specs_str)
                    sheet.cell(row=row, column=8, value=detailed_specs_str)
                    sheet.cell(row=row, column=9, value=attrs.get('whereToBuyLink', ''))
        
        temp_file = 'products_export.xlsx'
        workbook.save(temp_file)
        
        await update.callback_query.message.reply_document(
            document=open(temp_file, 'rb'),
            filename='products_export.xlsx',
            caption=f"✅ Выгружено {len(products)} товаров"
        )

        os.remove(temp_file)
        
        await update.callback_query.edit_message_text(
            "Экспорт завершен!\n\n"
            "Для начала нового экспорта или загрузки используйте /start"
        )
        
    except Exception as e:
        print(f"Error during export: {e}")
        await update.callback_query.edit_message_text(
            f"❌ Произошла ошибка при экспорте: {str(e)}\n\n"
            "Попробуйте еще раз используя /start"
        )

def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CallbackQueryHandler(button))
    application.add_handler(MessageHandler(filters.Document.ALL, process_excel))
    application.add_handler(MessageHandler(filters.TEXT, handle_message))

    print("Starting bot...")
    application.run_polling()

if __name__ == '__main__':
    main()

import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import json
import os
from urllib.parse import urljoin
from PIL import Image
from io import BytesIO
import concurrent.futures

def fetch_links(base_url, pages, link_class, log_widget, progress_var, progress_bar):
    product_links = []
    total_pages = pages
    for i in range(1, pages + 1):
        url = f"{base_url}/page/{i}/"
        try:
            response = requests.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            for a_tag in soup.find_all('a', class_=link_class):
                product_links.append(a_tag['href'])
            progress_var.set(i / total_pages * 100)
            progress_bar.update()
        except requests.exceptions.RequestException as e:
            log_widget.insert(tk.END, f"Failed to fetch page {i}: {e}\n")
        log_widget.see(tk.END)
    return list(set(product_links))  # Remove duplicates

def extract_data_single(link, search_terms, table_class, price_class, short_description_class, image_class, log_widget):
    page_data = {"link": link, "content": [], "table": [], "price": None, "short_description": [], "images": []}
    try:
        response = requests.get(link)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        for term in search_terms:
            elements = soup.find_all(class_=term)
            for element in elements:
                content = element.get_text(strip=True)
                page_data["content"].append(content)
                if term == search_terms[0]:  # Log only the first search term as the title
                    log_widget.insert(tk.END, f"Title found: {content}\n")

        table = soup.find("table", class_=table_class)
        if table:
            rows = table.find_all("tr")
            for row in rows:
                cells = [cell.get_text(strip=True) for cell in row.find_all(["th", "td"])]
                page_data["table"].append(cells)

        price_element = soup.find(class_=price_class)
        if price_element:
            del_element = price_element.find("del")
            if del_element:
                page_data["price"] = del_element.get_text(strip=True)

        short_description = soup.find(class_=short_description_class)
        if short_description:
            list_items = short_description.find_all("li")
            for item in list_items:
                page_data["short_description"].append(item.get_text(strip=True))

        images = soup.find_all(class_=image_class)
        for img in images:
            img_url = img.get("src")
            if img_url:
                img_url = urljoin(link, img_url)
                page_data["images"].append(img_url)
                image_response = requests.get(img_url)
                if image_response.status_code == 200:
                    if not os.path.exists("images"):
                        os.makedirs("images")
                    img_data = Image.open(BytesIO(image_response.content))
                    img_name = os.path.basename(img_url)
                    img_data.save(os.path.join("images", img_name))
    except requests.exceptions.RequestException as e:
        log_widget.insert(tk.END, f"Failed to fetch {link}: {e}\n")
    log_widget.see(tk.END)
    return page_data

def extract_data(links, search_terms, table_class, price_class, short_description_class, image_class, log_widget, progress_var, progress_bar):
    matches = []
    total_links = len(links)

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(extract_data_single, link, search_terms, table_class, price_class, short_description_class, image_class, log_widget)
            for link in links
        ]
        for idx, future in enumerate(concurrent.futures.as_completed(futures)):
            matches.append(future.result())
            progress_var.set((idx + 1) / total_links * 100)
            progress_bar.update()

    return matches

def save_to_excel(data, file_path, log_widget):
    records = []
    for match in data:
        records.append({
            "Link": match["link"],
            "Content": " | ".join(match["content"]),
            "Price": match["price"],
            "Short Description": " | ".join(match["short_description"]),
            "Table Data": " | ".join([" | ".join(row) for row in match["table"]]),
            "Images": " | ".join(match["images"])
        })
    df = pd.DataFrame(records)
    df.to_excel(file_path, index=False)
    log_widget.insert(tk.END, f"Data saved to {file_path}\n")
    log_widget.see(tk.END)

def save_configurations(config):
    configs = load_configurations()
    configs.append(config)
    with open('configurations.json', 'w') as f:
        json.dump(configs, f, indent=4)

def load_configurations():
    if os.path.exists('configurations.json'):
        with open('configurations.json', 'r') as f:
            return json.load(f)
    return []

def fetch_and_save(base_url, pages, link_class, search_terms, table_class, price_class, short_description_class, image_class, file_path, log_widget, progress_var, progress_bar):
    try:
        log_widget.insert(tk.END, "Fetching links...\n")
        links = fetch_links(base_url, pages, link_class, log_widget, progress_var, progress_bar)
        log_widget.insert(tk.END, "Extracting data...\n")
        data = extract_data(links, search_terms, table_class, price_class, short_description_class, image_class, log_widget, progress_var, progress_bar)
        log_widget.insert(tk.END, "Saving data to Excel...\n")
        save_to_excel(data, file_path, log_widget)
        messagebox.showinfo("Success", "Data has been saved successfully.")
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid inputs.")

def start_app():
    def on_fetch_and_save():
        base_url = url_entry.get()
        try:
            pages = int(pages_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid number of pages.")
            return
        link_class = link_class_entry.get().strip()
        search_terms = [term.strip() for term in search_terms_entry.get().split(',')]
        table_class = table_class_entry.get().strip()
        price_class = price_class_entry.get().strip()
        short_description_class = short_desc_class_entry.get().strip()
        image_class = image_class_entry.get().strip()
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        config = {
            "base_url": base_url,
            "pages": pages,
            "link_class": link_class,
            "search_terms": search_terms,
            "table_class": table_class,
            "price_class": price_class,
            "short_description_class": short_description_class,
            "image_class": image_class
        }
        save_configurations(config)
        threading.Thread(target=fetch_and_save, args=(
            base_url, pages, link_class, search_terms, table_class, price_class, short_description_class, image_class, file_path, log_widget, progress_var, progress_bar
        ), daemon=True).start()

    def on_load_config():
        configs = load_configurations()
        if configs:
            config_window = tk.Toplevel(root)
            config_window.title("Select Configuration")
            config_window.geometry("400x300")
            config_listbox = tk.Listbox(config_window, selectmode=tk.SINGLE, font=("Arial", 12))
            config_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            for idx, config in enumerate(configs):
                config_listbox.insert(tk.END, f"Config {idx + 1}: {config['base_url']}")

            def on_select_config():
                selection = config_listbox.curselection()
                if selection:
                    selected_config = configs[selection[0]]
                    url_entry.delete(0, tk.END)
                    url_entry.insert(0, selected_config["base_url"])
                    pages_entry.delete(0, tk.END)
                    pages_entry.insert(0, selected_config["pages"])
                    link_class_entry.delete(0, tk.END)
                    link_class_entry.insert(0, selected_config["link_class"])
                    search_terms_entry.delete(0, tk.END)
                    search_terms_entry.insert(0, ", ".join(selected_config["search_terms"]))
                    table_class_entry.delete(0, tk.END)
                    table_class_entry.insert(0, selected_config["table_class"])
                    price_class_entry.delete(0, tk.END)
                    price_class_entry.insert(0, selected_config["price_class"])
                    short_desc_class_entry.delete(0, tk.END)
                    short_desc_class_entry.insert(0, selected_config["short_description_class"])
                    image_class_entry.delete(0, tk.END)
                    image_class_entry.insert(0, selected_config["image_class"])
                    config_window.destroy()

            select_button = ttk.Button(config_window, text="Select", command=on_select_config)
            select_button.pack(pady=10)

    root = tk.Tk()
    root.title("Product Data Extractor")
    root.geometry("800x1000")
    root.configure(bg="#f0f0f0")

    style = ttk.Style()
    style.configure("TButton", font=("Arial", 12), padding=6)
    style.configure("TLabel", font=("Arial", 12), background="#f0f0f0")
    style.configure("TEntry", font=("Arial", 12))
    style.configure("TProgressbar", thickness=20)

    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    ttk.Label(main_frame, text="Base URL:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    url_entry = ttk.Entry(main_frame, width=60)
    url_entry.grid(row=0, column=1, padx=10, pady=5)
    url_entry.insert(0, "https://tajhizman.ir/product-category/heating-equipment")

    ttk.Label(main_frame, text="Number of Pages:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    pages_entry = ttk.Entry(main_frame, width=10)
    pages_entry.grid(row=1, column=1, padx=10, pady=5, sticky='w')
    pages_entry.insert(0, "15")

    ttk.Label(main_frame, text="Link Class:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
    link_class_entry = ttk.Entry(main_frame, width=60)
    link_class_entry.grid(row=2, column=1, padx=10, pady=5)
    link_class_entry.insert(0, "woocommerce-loop-product__link")

    ttk.Label(main_frame, text="Search Terms (comma separated):").grid(row=3, column=0, padx=10, pady=5, sticky='e')
    search_terms_entry = ttk.Entry(main_frame, width=60)
    search_terms_entry.grid(row=3, column=1, padx=10, pady=5)
    search_terms_entry.insert(0, "product_title entry-title, subtitle, woocommerce-Tabs-panel woocommerce-Tabs-panel--description panel entry-content wc-tab")

    ttk.Label(main_frame, text="Table Class:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
    table_class_entry = ttk.Entry(main_frame, width=60)
    table_class_entry.grid(row=4, column=1, padx=10, pady=5)
    table_class_entry.insert(0, "woocommerce-product-attributes shop_attributes")

    ttk.Label(main_frame, text="Price Class:").grid(row=5, column=0, padx=10, pady=5, sticky='e')
    price_class_entry = ttk.Entry(main_frame, width=60)
    price_class_entry.grid(row=5, column=1, padx=10, pady=5)
    price_class_entry.insert(0, "price-removed")

    ttk.Label(main_frame, text="Short Description Class:").grid(row=6, column=0, padx=10, pady=5, sticky='e')
    short_desc_class_entry = ttk.Entry(main_frame, width=60)
    short_desc_class_entry.grid(row=6, column=1, padx=10, pady=5)
    short_desc_class_entry.insert(0, "woocommerce-product-details__short-description")

    ttk.Label(main_frame, text="Image Class:").grid(row=7, column=0, padx=10, pady=5, sticky='e')
    image_class_entry = ttk.Entry(main_frame, width=60)
    image_class_entry.grid(row=7, column=1, padx=10, pady=5)
    image_class_entry.insert(0, "wp-post-image")

    fetch_button = ttk.Button(main_frame, text="Fetch and Save", command=on_fetch_and_save)
    fetch_button.grid(row=8, column=0, columnspan=2, pady=10)

    load_config_button = ttk.Button(main_frame, text="Load Saved Config", command=on_load_config)
    load_config_button.grid(row=9, column=0, columnspan=2, pady=5)

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(main_frame, variable=progress_var, maximum=100)
    progress_bar.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky='ew')

    log_widget = tk.Text(main_frame, height=15, wrap='word')
    log_widget.grid(row=11, column=0, columnspan=2, padx=10, pady=5)
    log_widget.configure(font=("Arial", 10))

    root.mainloop()

start_app()

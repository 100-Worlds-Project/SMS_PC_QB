import tkinter as tk
from tkinter import ttk, messagebox

def main_app():
    """
    Integrated app with user inputs at the top, dynamic results in a scrollable grid layout, and a close button.
    """
    def calculate_results():
        """
        Create a block for each medium, including its title, itemized costs, and separator line.
        """
        try:
            height = float(height_entry.get())
            width = float(width_entry.get())
            num_prints = int(num_prints_entry.get())

            # Clear previous results
            for widget in results_inner_frame.winfo_children():
                widget.destroy()
            
            # Bulk pricing data
            bulk_pricing = {
                "Canvas with Thick Gallery Wrap": [24.15, 22.05, 21.00, 19.95, 18.90],
                "Canvas with Basic Stretch": [24.15, 22.05, 21.00, 19.95, 18.90],
                "Canvas with Thin Gallery Wrap": [24.15, 22.05, 21.00, 19.95, 18.90],
                "Photorag": [22.05, 19.95, 18.90, 17.85, 16.80],
                "Enhanced Matte": [14.50, 12.50, 11.50, 10.50, 9.50],
                "Watercolor": [23.10, 21.00, 19.95, 18.90, 17.85],
                "Unstretched Canvas": [24.15, 22.05, 21.00, 19.95, 18.90],
            }

            frame_costs = {
                "Canvas with Thick Gallery Wrap": 4.99,
                "Canvas with Thin Gallery Wrap": 4.46,
                "Canvas with Basic Stretch": 4.46,
                "Default": 4.99,
            }

            gallery_stretching_fees = {
                "XS": (0, 36, 31.50),
                "S": (36, 56, 42.00),
                "M": (56, 72, 52.50),
                "L": (72, 100, 63.00),
                "XL": (100, float("inf"), 73.50),
            }

            basic_stretching_fees = {
                "XS": (0, 36, 21.00),
                "S": (36, 56, 31.50),
                "M": (56, 72, 42.00),
                "L": (72, 100, 52.50),
                "XL": (100, float("inf"), 63.00),
            }

            bracer_bar_cost = 4.46
            upcharge_72 = 150.00

            row = 0
            color_index = 0

            for print_type, prices in bulk_pricing.items():
                if num_prints <= 4:
                    price_per_sqft = prices[0]
                    pro_price_per_sqft = prices[1]
                elif num_prints <= 19:
                    price_per_sqft = prices[1]
                    pro_price_per_sqft = prices[2]
                elif num_prints <= 49:
                    price_per_sqft = prices[2]
                    pro_price_per_sqft = prices[3]
                elif num_prints <= 99:
                    price_per_sqft = prices[3]
                    pro_price_per_sqft = prices[4]
                else:
                    price_per_sqft = prices[4]
                    pro_price_per_sqft = (prices[4] * (prices[4] / prices[3]))

                gallery_wrap_adjustment = 3 if print_type.startswith("Canvas with T") else 0
                adjusted_height = height + gallery_wrap_adjustment
                adjusted_width = width + gallery_wrap_adjustment
                area_in_sqft = (adjusted_height * adjusted_width) / 144
                perimeter_in_feet = (2 * height + 2 * width) / 12

                if print_type.startswith("Canvas with"):
                    frame_cost = perimeter_in_feet * frame_costs.get(print_type, frame_costs["Default"])
                    bracer_cost = 0
                    if max(height, width) >= 40:
                        bracer_cost = (min(height, width) / 12) * bracer_bar_cost
                    if max(height, width) >= 60:
                        bracer_cost *= 2
                else:
                    frame_cost = 0
                    bracer_cost = 0
                    
                if print_type.startswith("Canvas with"):
                    upcharge = upcharge_72 if max(height, width) >= 72 else 0
                else:
                    upcharge = 0

                # Determine the stretching fee dictionary based on the print type
                if print_type.startswith("Canvas with T"):
                    stretching_fees = gallery_stretching_fees
                elif print_type == "Canvas with Basic Stretch":
                    stretching_fees = basic_stretching_fees
                else:
                    stretching_fees = {}

                # Calculate the stretching fee
                stretching_fee = 0
                for category, (min_range, max_range, fee) in stretching_fees.items():
                    if min_range <= height + width <= max_range:
                        stretching_fee = fee
                        break
        
                canvas_cost = area_in_sqft * price_per_sqft
                pro_canvas_cost = area_in_sqft * pro_price_per_sqft
                total_cost = canvas_cost + frame_cost + bracer_cost + upcharge + stretching_fee
                pro_total_cost = pro_canvas_cost + frame_cost + bracer_cost + upcharge + stretching_fee

                # Define a list of colors to cycle through
                colors = ["#ccff00", "#00ccff", "#ff00cc", "#ffcc00", "#cc00ff", "#00ffcc", "#ff0000"]
                
                # Cycle through the colors
                current_color = colors[color_index % len(colors)]

                block_text = f"{print_type}\n"
                block_text += f"    Print:{'.' * (30 - len('    Print:'))} $ {canvas_cost:.2f}\n"
                block_text += f"    Print (Pro Discount):{'.' * (30 - len('    Print (Pro Discount):'))} $ {pro_canvas_cost:.2f}\n"
                block_text += f"    Frame Wood:{'.' * (30 - len('    Frame Wood:'))} $ {frame_cost:.2f}\n"
                block_text += f"    Stretch:{'.' * (30 - len('    Stretch:'))} $ {stretching_fee:.2f}\n"
                block_text += f"    Bracer Wood:{'.' * (30 - len('    Bracer Wood:'))} $ {bracer_cost:.2f}\n"
                block_text += f"    ≥ 72\" Upcharge:{'.' * (30 - len('    ≥ 72\" Upcharge:'))} $ {upcharge:.2f}\n"
                block_text += f"    Regular Price:{'.' * (30 - len('    Regular Price:'))} $ {total_cost:.2f}\n"
                block_text += f"    Professional Price:{'.' * (30 - len('    Professional Price:'))} $ {pro_total_cost:.2f}\n"
                block_text += "=" * 45 + "\n"

                # Create the Text widget
                text_widget = tk.Text(
                    results_inner_frame,
                    height=len(block_text.splitlines()),
                    width=80,
                    font=("Monaco", 12),
                    bg="#333333",
                    fg="#e0e0e0",
                    bd=0,
                    wrap="none"
                )
                text_widget.insert("1.0", block_text)

                # Highlight the title (first line)
                text_widget.tag_configure("title", font=("Arial", 15, "bold"), foreground=current_color)
                text_widget.tag_add("title", "1.0", "2.0")

                # Highlight the separator line (last line in the block)
                separator_line_start = len(block_text.splitlines())
                text_widget.tag_configure("separator", font=("Arial", 15, "bold"), foreground=current_color)
                text_widget.tag_add("separator", f"{separator_line_start}.0", f"{separator_line_start}.end")

                # Highlight the totals (bold white for the last two lines before the separator)
                total_start_line = separator_line_start - 2  # Two lines above the separator
                text_widget.tag_configure("highlight", font=("Monaco", 12, "bold"), foreground="#ffffff")
                text_widget.tag_add("highlight", f"{total_start_line}.0", f"{separator_line_start}.0")

                text_widget.configure(state="disabled")
                text_widget.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=10)


                color_index += 1

                # Return the updated row count
                row += len(block_text.splitlines())

                results_inner_frame.update_idletasks()
                results_canvas.config(scrollregion=results_canvas.bbox("all"))

        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers for all fields.")

    def exit_program(event=None):
        root.quit()
        root.destroy()

    def scroll_mac(event):
        results_canvas.yview_scroll(-1 * int(event.delta), "units")

    root = tk.Tk()
    root.withdraw()  # Hide the weird default root window

    app = tk.Toplevel(root)
    app.title("Print Cost Calculator")
    app.geometry("1500x1500")

    # Input frame
    input_frame = tk.Frame(app)
    input_frame.pack(fill="x", padx=10, pady=10)

    tk.Label(input_frame, text="Height (inches):", font=("Arial", 15, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="e")
    height_entry = tk.Entry(input_frame, font=("Arial", 15, "bold"))
    height_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(input_frame, text="Width (inches):", font=("Arial", 15, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="e")
    width_entry = tk.Entry(input_frame, font=("Arial", 15, "bold"))
    width_entry.grid(row=0, column=3, padx=5, pady=5)

    tk.Label(input_frame, text="Number of Prints:", font=("Arial", 15, "bold")).grid(row=0, column=4, padx=5, pady=5, sticky="e")
    num_prints_entry = tk.Entry(input_frame, font=("Arial", 15, "bold"))
    num_prints_entry.grid(row=0, column=5, padx=5, pady=5)

    tk.Button(input_frame, text="Calculate", font=("Arial", 12, "bold"), command=calculate_results).grid(row=0, column=6, padx=10, pady=5)

    # Results frame with scrollable canvas
    results_frame = tk.Frame(app)
    results_frame.pack(fill="both", expand=True, padx=10, pady=10)

    results_canvas = tk.Canvas(results_frame)
    scrollbar_y = ttk.Scrollbar(results_frame, orient="vertical", command=results_canvas.yview)
    results_inner_frame = tk.Frame(results_canvas)

    results_canvas.create_window((0, 0), window=results_inner_frame, anchor="nw")
    results_canvas.configure(yscrollcommand=scrollbar_y.set)
    results_canvas.pack(side="left", fill="both", expand=True)
    scrollbar_y.pack(side="right", fill="y")

    # Bind scrolling to the canvas
    app.bind("<MouseWheel>", scroll_mac)

    close_button = tk.Button(app, text="Close", font=("Arial", 12, "bold"), command=exit_program)
    close_button.pack(pady=10)

    # Bind Enter key to submit inputs
    app.bind("<Return>", lambda event: calculate_results())

    # Bind Escape key to exit program
    app.bind("<Escape>", exit_program)

        # Bind Up and Down arrow keys to scroll functionality
    app.bind("<Up>", lambda event: results_canvas.yview_scroll(-1, "units"))
    app.bind("<Down>", lambda event: results_canvas.yview_scroll(1, "units"))

    height_entry.focus()

    root.mainloop()


# Run the main app
main_app()

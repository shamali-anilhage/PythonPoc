from flask import Flask, request, jsonify
import openpyxl

app = Flask(__name__)

# Define the route to handle the HTTP request
@app.route('/add-row', methods=['POST'])
def add_row():
    try:
        # Get data from the incoming HTTP request body (expects JSON format)
        data = request.get_json()

        # Print the data to check its structure
        print(f"Received data: {data}")

        # Load the existing Excel workbook (make sure the path is correct)
        wb = openpyxl.load_workbook('C:/Users/sanilhage/Downloads/PaymentUpdateOutputTemplate (1).xlsx')
        sheet = wb.active

        # Check if the data is a list of dictionaries
        if isinstance(data, list):
            for row in data:
                if isinstance(row, dict):
                    new_row = list(row.values())  # Extract values from the dictionary
                    sheet.append(new_row)  # Append the new row to the Excel sheet
                else:
                    raise ValueError("Each item in the list should be a dictionary.")
        else:
            raise ValueError("Expected a list of dictionaries.")

        # Save the workbook with the new rows
        wb.save('C:/Users/sanilhage/Downloads/PaymentUpdateOutputTemplate (1).xlsx')

        return jsonify({"message": "Rows added successfully"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # Run the Flask app (make sure to bind to the correct host and port)
    app.run(host='0.0.0.0', port=5000)

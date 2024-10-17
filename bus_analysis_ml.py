from flask import Flask, request, jsonify
import xlsxwriter
import smtplib
from email.mime.text import MIMEText

app = Flask(__name__)

def calculate_efficiency(vehicles):
    for vehicle in vehicles:
        if vehicle["Diesel_consume"] > 0:  
            vehicle["Efficiency"] = vehicle["Distance"] / vehicle["Diesel_consume"]
        else:
            vehicle["Efficiency"] = 0 
    return vehicles

def write_to_excel(vehicles):
    workbook = xlsxwriter.Workbook("VehicleData.xlsx")
    worksheet = workbook.add_worksheet("Vehicle Details")
    
    headers = ["Vehicle No", "Distance (km)", "Diesel Consumed (liters)", "Time Taken (hours)", "Efficiency (km/l)", "Ticket Sales ($)"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    for index, vehicle in enumerate(vehicles):
        worksheet.write(index + 1, 0, vehicle["Vehicle_no"])
        worksheet.write(index + 1, 1, vehicle["Distance"])
        worksheet.write(index + 1, 2, vehicle["Diesel_consume"])
        worksheet.write(index + 1, 3, vehicle["Time_taken"])
        worksheet.write(index + 1, 4, vehicle["Efficiency"])
        worksheet.write(index + 1, 5, vehicle["Ticket_sales"])

    workbook.close()

def send_email(best_vehicle, worst_vehicle, total_ticket_sales):
    sender_email = "sender email" 
    receiver_email = "recever email"
    password = "password of the sender email"  

    subject = "Vehicle Efficiency Report"
    body = (
        f"The best vehicle is {best_vehicle['Vehicle_no']} with efficiency {best_vehicle['Efficiency']:.2f} km/l.\n"
        f"The worst vehicle is {worst_vehicle['Vehicle_no']} with efficiency {worst_vehicle['Efficiency']:.2f} km/l.\n"
        f"Total ticket sales amount to: ${total_ticket_sales:.2f}."
    )

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return "Email sent successfully."
    except Exception as e:
        return f"Failed to send email: {e}"

@app.route('/generate_report', methods=['POST'])
def generate_report():
    data = request.json
    vehicles = data['vehicles']

    # Rename keys to match the original script
    for vehicle in vehicles:
        vehicle['Vehicle_no'] = vehicle.pop('vehicleNo')
        vehicle['Diesel_consume'] = vehicle.pop('dieselConsumed')
        vehicle['Time_taken'] = vehicle.pop('timeTaken')
        vehicle['Ticket_sales'] = vehicle.pop('ticketSales')

    vehicles = calculate_efficiency(vehicles)
    write_to_excel(vehicles)

    best_vehicle = max(vehicles, key=lambda x: x["Efficiency"])
    worst_vehicle = min(vehicles, key=lambda x: x["Efficiency"])
    total_ticket_sales = sum(vehicle["Ticket_sales"] for vehicle in vehicles)

    email_result = send_email(best_vehicle, worst_vehicle, total_ticket_sales)

    return jsonify({
        'best_vehicle': best_vehicle,
        'worst_vehicle': worst_vehicle,
        'total_ticket_sales': total_ticket_sales,
        'email_result': email_result,
        'vehicles': vehicles
    })

if __name__ == '__main__':
    app.run(debug=True)

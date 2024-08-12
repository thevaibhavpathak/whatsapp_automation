To use this script effectively, your Excel file should have the following structure:

| **Column A**            | **Column B** |
| ----------------------- | ------------ |
| **Contact Name/Number** | **Message**  |

### **Details:**

1. **Column A - Contact Name/Number:**

   - **Contact Name:** This should be the exact name of the contact as it appears in your WhatsApp contacts. For instance, if the contact is saved as "John Doe" in WhatsApp, the corresponding cell should contain "John Doe".
   - **Phone Number:** Alternatively, you can use the contact's phone number. It should be in the international format (e.g., +1XXXXXXXXXX for the USA, +91XXXXXXXXXX for India).

2. **Column B - Message:**
   - This column should contain the message you wish to send to each respective contact.
   - Messages can include text, emojis, and even URLs. Ensure that each message corresponds to the correct contact or phone number listed in Column A.

### **Example Excel File:**

| **Contact Name/Number** | **Message**                             |
| ----------------------- | --------------------------------------- |
| John Doe                | Hey John, just checking in!             |
| +1234567890             | Hi there! Don't forget the meeting.     |
| Jane Smith              | Your package has been shipped.          |
| +9876543210             | Happy Birthday! 🎉                      |
| Group Chat Name         | Reminder: Project deadline is tomorrow. |

### **Usage:**

- The script reads the data from the Excel file, starting from the second row (A2 and B2), and sends the corresponding message to the contact or phone number in each row.
- Ensure there are no empty rows between your data entries, as this could disrupt the script's operation.

### **Saving the Excel File:**

- Save the Excel file as .xlsx or .xlsm if you're using it within the same workbook as the script.
- Make sure the file is located in a known directory so that it can be easily referenced when running the script.

This structure allows for efficient and accurate message sending, ensuring each contact receives the correct message.

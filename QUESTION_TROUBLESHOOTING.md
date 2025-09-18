# 🔍 Question Function Troubleshooting Guide

## Issue: Questions failing in Excel-Ollama plugin

The question function might be failing due to several reasons. Let's diagnose and fix them step by step.

## Step 1: Use the Debug Version

Replace your current VBA code with the debug version that has enhanced error reporting:

### Installation:
1. **Open Excel** → Press **Alt+F11** (VBA Editor)
2. **Find your add-in project** and delete the old module
3. **Insert → Module** (create new one)
4. **Copy ALL content** from `debug-excel-addin.vba`
5. **Update YOUR_EC2_IP** with your actual EC2 IP address
6. **Save** and close VBA Editor

## Step 2: Run Diagnostic Tests

### Test 1: Basic Connection
1. **Press Alt+F8**
2. **Run:** `TestBasicConnection`
3. **Expected:** Should show "Connection successful" with JSON response

### Test 2: Simple Generation
1. **Press Alt+F8**
2. **Run:** `TestSimpleGeneration`
3. **Expected:** Should get a simple response from Ollama

### Test 3: Enhanced Question
1. **Select some data** in Excel (including headers)
2. **Press Alt+F8**
3. **Run:** `AskQuestionAboutDataEnhanced`
4. **Ask a simple question:** "What is this data about?"
5. **Check the result sheet** for detailed debug information

## Common Issues and Solutions:

### Issue 1: JSON Parsing Errors

**Symptoms:** Error messages about JSON parsing
**Solution:** The debug version has improved JSON handling

### Issue 2: Prompt Too Long

**Symptoms:** Request fails with large datasets
**Solution:** The debug version limits sample data to 3 rows

### Issue 3: Special Characters in Data

**Symptoms:** JSON errors with certain data
**Solution:** Enhanced JSON escaping in debug version

### Issue 4: Model Not Responding

**Symptoms:** Long delays or timeouts
**Solution:** 
```bash
# On EC2, check if model is loaded
ssh -i your-key.pem ec2-user@YOUR_EC2_IP
ollama list
ollama run llama2:latest  # Pre-load the model
```

### Issue 5: Server Configuration

**Symptoms:** Connection errors
**Solution:**
1. **Run:** `ConfigureOllamaServer`
2. **Set URL:** `http://YOUR_EC2_IP:11434`
3. **Set Model:** `llama2:latest`
4. **Test:** `TestBasicConnection`

## Debug Information Analysis:

When you run `AskQuestionAboutDataEnhanced`, check the result sheet for:

### Good Signs:
- ✅ HTTP Status: 200
- ✅ Response length: > 0 characters
- ✅ Successfully extracted response

### Bad Signs:
- ❌ HTTP Status: 404, 500, etc.
- ❌ Response length: 0 characters
- ❌ Connection errors
- ❌ JSON parsing errors

## Manual Testing:

### Test Ollama Directly:
```bash
# From your Windows laptop
curl -X POST http://YOUR_EC2_IP:11434/api/generate ^
  -H "Content-Type: application/json" ^
  -d "{\"model\":\"llama2:latest\",\"prompt\":\"Hello, please respond with 'Test successful!'\",\"stream\":false}"
```

### Expected Response:
```json
{
  "model": "llama2:latest",
  "created_at": "2024-01-01T00:00:00Z",
  "response": "Test successful!",
  "done": true
}
```

## Quick Fixes:

### Fix 1: Simplify Questions
Instead of complex questions, try:
- "What columns are in this data?"
- "How many rows are there?"
- "What is the first value?"

### Fix 2: Reduce Data Size
- Select smaller data ranges (< 100 rows)
- Use simple column names (no special characters)
- Ensure clean data (no empty cells in headers)

### Fix 3: Check Model Status
```bash
# On EC2
ollama ps  # Check if model is running
ollama run llama2:latest  # Start interactive session to pre-load
# Type /bye to exit
```

### Fix 4: Restart Ollama
```bash
# On EC2
sudo systemctl restart ollama
sleep 10
curl http://localhost:11434/api/tags
```

## Sample Test Data:

Create this simple test data in Excel:

| Name | Age | City |
|------|-----|------|
| John | 25  | NYC  |
| Jane | 30  | LA   |
| Bob  | 35  | Chicago |

Then ask: "What is the average age?"

## Error Message Decoder:

| Error Message | Likely Cause | Solution |
|---------------|--------------|----------|
| "Connection Error" | Network/Server issue | Check EC2 status, Security Group |
| "HTTP Error 404" | Wrong URL | Verify server URL configuration |
| "HTTP Error 500" | Server error | Check Ollama logs, restart service |
| "No response field found" | JSON parsing issue | Use debug version |
| "Could not parse response" | Malformed JSON | Check Ollama model status |

## Advanced Debugging:

### Check Ollama Logs:
```bash
# On EC2
sudo journalctl -u ollama -f
```

### Check Network:
```bash
# On EC2
sudo netstat -tlnp | grep 11434
curl http://localhost:11434/api/tags
```

### Test with Different Models:
```bash
# On EC2
ollama pull mistral:latest
# Then update Excel plugin to use mistral:latest
```

The debug version will give you detailed information about what's failing. Run the diagnostic tests and check the debug output to identify the specific issue!
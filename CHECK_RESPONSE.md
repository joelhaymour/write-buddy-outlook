# Check the 400 Error Response

## Important: We Need the Response Body

The 400 Bad Request error has a response body that tells us exactly what's wrong. Please check:

### Steps:
1. In Chrome DevTools, go to the **Network** tab
2. Find the failed request: `POST https://titles.prod.mos.microsoft.com/dev/v1/users/packages/addins`
3. **Click on that request**
4. Go to the **Response** tab (or **Preview** tab)
5. **Copy the entire response** - it should tell us what validation error Microsoft found

The response will look something like:
```json
{
  "error": {
    "code": "...",
    "message": "The manifest is invalid because..."
  }
}
```

Or it might be XML or plain text explaining the validation error.

## Common Issues That Cause 400:
- Missing required elements
- Invalid element values
- Schema validation failures
- Invalid URL formats
- Missing resource references

**Please share the response body** so we can fix the exact issue!


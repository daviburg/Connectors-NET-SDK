//------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
//------------------------------------------------------------

using System.Collections.Generic;
using System.Text.Json;
using Azure.Connectors.Sdk.Office365;
using Azure.Connectors.Sdk.Office365.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Azure.Connectors.Sdk.Tests
{
    /// <summary>
    /// Tests for trigger callback payload deserialization.
    /// </summary>
    [TestClass]
    public class TriggerCallbackPayloadTests
    {
        /// <summary>
        /// Captured Connector Gateway trigger callback payload from production (2026-03-16).
        /// </summary>
        private const string CapturedTriggerPayload = """
            {
              "body": {
                "value": [
                  {
                    "id": "AAMkADlmOTA3NWNm",
                    "receivedDateTime": "2026-03-16T21:26:21+00:00",
                    "hasAttachments": false,
                    "subject": "Test email for trigger callback",
                    "bodyPreview": "This is a test email preview.",
                    "importance": "normal",
                    "isRead": false,
                    "isHtml": true,
                    "body": "<html><body>Hello</body></html>",
                    "from": "sender@microsoft.com",
                    "toRecipients": "recipient1@microsoft.com;recipient2@microsoft.com",
                    "ccRecipients": "cc@microsoft.com",
                    "bccRecipients": null,
                    "replyTo": null,
                    "attachments": []
                  }
                ]
              }
            }
            """;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        [TestMethod]
        public void Deserialize_CapturedPayload_ReturnsEnvelopeWithOneEmail()
        {
            // Arrange & Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedTriggerPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Body);
            Assert.IsNotNull(result.Body.Value);
            Assert.AreEqual(1, result.Body.Value.Count);
        }

        [TestMethod]
        public void Deserialize_CapturedPayload_ParsesEmailFields()
        {
            // Arrange & Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedTriggerPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Body);
            Assert.IsNotNull(result.Body.Value);
            Assert.IsTrue(result.Body.Value.Count > 0, "Expected at least one email in the value array.");
            var email = result.Body.Value[0];

            // Assert
            Assert.AreEqual("AAMkADlmOTA3NWNm", email.MessageId);
            Assert.AreEqual("Test email for trigger callback", email.Subject);
            Assert.AreEqual("sender@microsoft.com", email.From);
            Assert.AreEqual("recipient1@microsoft.com;recipient2@microsoft.com", email.To);
            Assert.AreEqual("cc@microsoft.com", email.CC);
            Assert.AreEqual("normal", email.Importance);
            Assert.AreEqual("This is a test email preview.", email.BodyPreview);
            Assert.AreEqual(false, email.HasAttachment);
            Assert.AreEqual(false, email.IsRead);
            Assert.AreEqual(true, email.IsHTML);
            Assert.AreEqual("<html><body>Hello</body></html>", email.Body);
        }

        [TestMethod]
        public void Deserialize_MultipleEmails_ParsesAll()
        {
            // Arrange
            var multiPayload = """
                {
                  "body": {
                    "value": [
                      { "subject": "Email 1", "from": "a@test.com" },
                      { "subject": "Email 2", "from": "b@test.com" },
                      { "subject": "Email 3", "from": "c@test.com" }
                    ]
                  }
                }
                """;

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                multiPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.AreEqual(3, result!.Body!.Value!.Count);
            Assert.AreEqual("Email 1", result.Body.Value[0].Subject);
            Assert.AreEqual("Email 3", result.Body.Value[2].Subject);
        }

        [TestMethod]
        public void Deserialize_EmptyValueArray_ReturnsEmptyList()
        {
            // Arrange
            var emptyPayload = """{"body":{"value":[]}}""";

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                emptyPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.IsNotNull(result!.Body!.Value);
            Assert.AreEqual(0, result.Body.Value.Count);
        }

        [TestMethod]
        public void Deserialize_NullBody_ReturnsNullBody()
        {
            // Arrange
            var nullBodyPayload = """{"body":null}""";

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                nullBodyPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.IsNull(result!.Body);
        }

        [TestMethod]
        public void Deserialize_WithAttachments_ParsesAttachmentMetadata()
        {
            // Arrange
            var payload = """
                {
                  "body": {
                    "value": [
                      {
                        "subject": "Email with attachment",
                        "hasAttachments": true,
                        "attachments": [
                          {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": "document.pdf",
                            "contentType": "application/pdf",
                            "size": 12345
                          }
                        ]
                      }
                    ]
                  }
                }
                """;

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                payload,
                TriggerCallbackPayloadTests.JsonOptions);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Body);
            Assert.IsNotNull(result.Body.Value);
            Assert.IsTrue(result.Body.Value.Count > 0, "Expected at least one email in the value array.");
            var email = result.Body.Value[0];

            // Assert
            Assert.AreEqual(true, email.HasAttachment);
            Assert.IsNotNull(email.Attachments);
            Assert.AreEqual(1, email.Attachments.Count);
            Assert.AreEqual("document.pdf", email.Attachments[0].Name);
            Assert.AreEqual("application/pdf", email.Attachments[0].ContentType);
            Assert.AreEqual(12345L, email.Attachments[0].Size);
        }

        [TestMethod]
        public void Deserialize_GenericType_WorksWithDictionary()
        {
            // Arrange — demonstrate that the envelope works with arbitrary types
            var payload = """{"body":{"value":[{"key":"value1"},{"key":"value2"}]}}""";

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<Dictionary<string, string>>>(
                payload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.AreEqual(2, result!.Body!.Value!.Count);
            Assert.AreEqual("value1", result.Body.Value[0]["key"]);
            Assert.AreEqual("value2", result.Body.Value[1]["key"]);
        }

        #region Single-item payload tests (GitHub issue #149)

        /// <summary>
        /// Captured Connector Namespace trigger callback payload for OnNewEmailV3 (2026-05-14).
        /// This trigger delivers a single email object directly in body — no "value" array wrapper.
        /// See: https://github.com/Azure/Connectors-NET-SDK/issues/149
        /// </summary>
        private const string CapturedSingleItemPayload = """
            {
              "body": {
                "id": "AAMkADQ0MTI1NTBm",
                "receivedDateTime": "2026-05-14T13:53:19-07:00",
                "hasAttachments": false,
                "subject": "Test single-item trigger",
                "bodyPreview": "This email was delivered without a value array wrapper.",
                "importance": "normal",
                "isRead": false,
                "isHtml": true,
                "body": "<html><body>Hello from OnNewEmailV3</body></html>",
                "from": "thiago@microsoft.com",
                "toRecipients": "recipient@microsoft.com",
                "ccRecipients": null,
                "bccRecipients": null,
                "replyTo": null,
                "attachments": []
              }
            }
            """;

        [TestMethod]
        public void Deserialize_SingleItemPayload_NormalizesToValueList()
        {
            // Act — single-item shape: body IS the email, not {"value": [...]}
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedSingleItemPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert — converter normalizes single item into Value list
            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Body);
            Assert.IsNotNull(result.Body.Value);
            Assert.AreEqual(1, result.Body.Value.Count);
        }

        [TestMethod]
        public void Deserialize_SingleItemPayload_ParsesEmailFields()
        {
            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedSingleItemPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            var email = result!.Body!.Value![0];

            // Assert
            Assert.AreEqual("AAMkADQ0MTI1NTBm", email.MessageId);
            Assert.AreEqual("Test single-item trigger", email.Subject);
            Assert.AreEqual("thiago@microsoft.com", email.From);
            Assert.AreEqual("recipient@microsoft.com", email.To);
            Assert.AreEqual("normal", email.Importance);
            Assert.AreEqual(false, email.HasAttachment);
            Assert.AreEqual(false, email.IsRead);
            Assert.AreEqual(true, email.IsHTML);
        }

        [TestMethod]
        public void Deserialize_SingleItemPayload_BatchPayloadsStillWork()
        {
            // Arrange — batch shape should continue to work with the converter
            var batchResult = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedTriggerPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            var singleResult = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedSingleItemPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert — both produce non-null Value lists with 1 email
            Assert.AreEqual(1, batchResult!.Body!.Value!.Count);
            Assert.AreEqual(1, singleResult!.Body!.Value!.Count);

            // Assert — both yield the expected subjects
            Assert.AreEqual("Test email for trigger callback", batchResult.Body.Value[0].Subject);
            Assert.AreEqual("Test single-item trigger", singleResult.Body.Value[0].Subject);
        }

        [TestMethod]
        public void Deserialize_SingleItemPayload_GenericDictionary()
        {
            // Arrange — single-item shape with Dictionary<string, string>
            var payload = """{"body":{"key1":"val1","key2":"val2"}}""";

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<Dictionary<string, string>>>(
                payload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert — no "value" array → whole body deserialized as one Dictionary item
            Assert.IsNotNull(result!.Body!.Value);
            Assert.AreEqual(1, result.Body.Value.Count);
            Assert.AreEqual("val1", result.Body.Value[0]["key1"]);
            Assert.AreEqual("val2", result.Body.Value[0]["key2"]);
        }

        [TestMethod]
        public void Deserialize_SingleItemPayload_NullBody_StillReturnsNull()
        {
            // Arrange
            var payload = """{"body":null}""";

            // Act
            var result = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                payload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert — null body is preserved
            Assert.IsNull(result!.Body);
        }

        [TestMethod]
        public void Serialize_RoundTrip_BatchPayload()
        {
            // Arrange
            var original = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedTriggerPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Act — serialize and deserialize again
            var json = JsonSerializer.Serialize(original, TriggerCallbackPayloadTests.JsonOptions);
            var roundTripped = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                json,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.AreEqual(1, roundTripped!.Body!.Value!.Count);
            Assert.AreEqual("Test email for trigger callback", roundTripped.Body.Value[0].Subject);
        }

        [TestMethod]
        public void Serialize_RoundTrip_SingleItemPayload()
        {
            // Arrange — deserialize single-item payload
            var original = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                TriggerCallbackPayloadTests.CapturedSingleItemPayload,
                TriggerCallbackPayloadTests.JsonOptions);

            // Act — serialize (always outputs batch shape) and deserialize again
            var json = JsonSerializer.Serialize(original, TriggerCallbackPayloadTests.JsonOptions);
            var roundTripped = JsonSerializer.Deserialize<TriggerCallbackPayload<GraphClientReceiveMessage>>(
                json,
                TriggerCallbackPayloadTests.JsonOptions);

            // Assert
            Assert.AreEqual(1, roundTripped!.Body!.Value!.Count);
            Assert.AreEqual("Test single-item trigger", roundTripped.Body.Value[0].Subject);
        }

        #endregion Single-item payload tests (GitHub issue #149)
    }
}

# MCP Server Specification: Teams Collaboration Gaps

**Date**: 2024-01-24
**Context**: Identified during Synthetic Employees NPC Lifecycle System testing
**Priority**: High - Blocks true multi-agent collaboration workflows

---

## Executive Summary

During end-to-end testing of a multi-agent workflow (KAM delegating work to Marketing, Proofreader, and Editor), we discovered critical gaps in Teams functionality. The agents were forced to fall back to email-only communication, missing the opportunity for **true real-time collaboration**.

### The Vision We're Missing

When a project comes in, the NPC agents should:

1. **Create a dedicated Teams Channel** for the project
2. **Schedule a kickoff meeting** with all team members
3. **Share project documents** in the channel's Files tab
4. **Collaborate in real-time** - chatting, sharing findings, @mentioning
5. **Researchers post .md files** with their findings for others to consume
6. **Invite the client** (future) as a guest to the channel

Instead, we had to fall back to disconnected emails - no shared workspace, no real-time visibility, no collaborative document space.

---

## The Collaboration Model We Need

```
┌─────────────────────────────────────────────────────────────────────────┐
│                     PROJECT TEAMS CHANNEL                                │
│                     "Prysmian Submarine Cable - Project 41e06072"       │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                          │
│  📁 Files Tab                                                            │
│  ├── project-brief.md (KAM uploads client requirements)                  │
│  ├── research-prysmian.md (Researcher's findings)                        │
│  ├── draft-v1.md (Marketing Copywriter's first draft)                    │
│  ├── proofreading-notes.md (Proofreader's review)                        │
│  └── final-deliverable.md (Approved content)                             │
│                                                                          │
│  💬 Conversations                                                        │
│  ├── [Anna] Created project channel. Kickoff meeting scheduled for 2pm   │
│  ├── [Christina] Found great info about Prysmian's recent acquisition    │
│  │   └── [François] @Christina can you add that to research-prysmian.md? │
│  ├── [Bruno] Draft looks good, minor spelling fixes noted                │
│  │   └── [Christina] Fixed! Updated draft-v1.md                          │
│  ├── [François] Editorial review complete. Approved ✓                    │
│  └── [Anna] Sending to client now. Great work team!                      │
│                                                                          │
│  👥 Members (added by KAM when channel created)                           │
│  ├── Anna Kowalski (KAM) - Owner [creator]                               │
│  ├── Christina Hall (Marketing) [added via addChannelMember]             │
│  ├── Bruno Dupont (Proofreader) [added via addChannelMember]             │
│  ├── François Moreau (Editor) [added via addChannelMember]               │
│  └── Marcus Jensen (Client - Guest) [FUTURE: inviteGuestToTeam]          │
│                                                                          │
└─────────────────────────────────────────────────────────────────────────┘
```

This is TRUE NPC collaboration - a shared workspace where agents work together visibly.

---

## CRITICAL GAPS: Project Collaboration Workspace

These gaps prevent the NPC collaboration model described above.

### Gap A: Create Project Channel

**Current State**: No tool to create a new channel in a Team
**Impact**: Cannot create dedicated workspace for each project

**Required Tool**:
```
Tool Name: createTeamChannel
Description: Create a new channel in a Team for project collaboration
Parameters:
  - teamId (required): The Team ID
  - displayName (required): Channel name (e.g., "Project-41e06072-Prysmian")
  - description (optional): Channel description
  - membershipType (optional): "standard" or "private" (default: "standard")
Returns: Channel object with id, displayName, webUrl
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-post
- Endpoint: `POST /teams/{team-id}/channels`

**Request Body**:
```json
{
  "displayName": "Project-41e06072-Prysmian",
  "description": "Marketing flyer project for Nordic Subsea Infrastructure",
  "membershipType": "standard"
}
```

---

### Gap B: Add Members to Channel

**Current State**: No tool to add team members to a channel
**Impact**: Cannot assemble the project team in the channel - they can't see it or participate!

**Why This Is Critical**:
When KAM creates a project channel, the assigned team members (Writer, Proofreader, Editor) need to be **added to the channel** before they can:
- See the channel in their Teams
- Read files uploaded there
- Participate in chat
- Receive notifications

**Required Tool**:
```
Tool Name: addChannelMember
Description: Add a user to a Teams channel
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - userEmail (required): Email address of user to add
  - roles (optional): Array of roles ["owner"] or [] for member (default: member)
Returns: Membership confirmation with user details
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-post-members
- Endpoint: `POST /teams/{team-id}/channels/{channel-id}/members`

**Request Body**:
```json
{
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  "roles": [],
  "user@odata.bind": "https://graph.microsoft.com/v1.0/users('christina.hall@company.com')"
}
```

**Workflow Usage**:
```python
# After creating channel, add all project members:
channel = kam.create_channel("Project-41e06072")
kam.add_channel_member(channel, "christina.hall@...")   # Writer
kam.add_channel_member(channel, "bruno.dupont@...")     # Proofreader
kam.add_channel_member(channel, "francois.moreau@...")  # Editor
# Now they can all see and use the channel
```

**Note**: For standard channels in Teams, members of the Team can see all standard channels. For **private channels**, members must be explicitly added (which is what this tool does). Either way, having this tool ensures the right people are in the project workspace.

---

### Gap C: Upload File to Channel

**Current State**: File upload exists but not tied to channel's SharePoint folder
**Impact**: Researchers cannot share .md files in the project channel

**Required Tool**:
```
Tool Name: uploadFileToChannel
Description: Upload a file to a channel's Files tab (SharePoint folder)
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - fileName (required): Name of the file (e.g., "research-findings.md")
  - content (required): File content (text or base64 for binary)
  - contentType (optional): MIME type (default: inferred from extension)
Returns: DriveItem object with id, webUrl, downloadUrl
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
- The channel's files are in SharePoint: `/teams/{team-id}/channels/{channel-id}/filesFolder`
- Upload endpoint: `PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content`

**Workflow**:
1. Get channel's filesFolder driveItem
2. Upload file to that folder
3. File appears in channel's "Files" tab

---

### Gap D: List Files in Channel

**Current State**: No tool to list files in a channel's Files tab
**Impact**: NPCs cannot see what documents teammates have shared

**Required Tool**:
```
Tool Name: listChannelFiles
Description: List files in a channel's Files tab
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
Returns: Array of DriveItem objects with name, webUrl, lastModifiedBy
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-get-filesfolder
- https://learn.microsoft.com/en-us/graph/api/driveitem-list-children
- Endpoint: `GET /teams/{team-id}/channels/{channel-id}/filesFolder` then list children

---

### Gap E: Read File from Channel

**Current State**: General file read exists but not channel-aware
**Impact**: NPCs cannot read .md files shared by teammates in the channel

**Required Tool**:
```
Tool Name: readChannelFile
Description: Read content of a file from channel's Files tab
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - fileName (required): Name of the file to read
Returns: File content (text for .md/.txt, base64 for binary)
```

---

### Gap F: Invite External Guest to Channel (Future)

**Current State**: No guest invitation capability
**Impact**: Cannot include client in project channel for transparency

**Required Tool** (Future Enhancement):
```
Tool Name: inviteGuestToTeam
Description: Invite external user as guest to Team
Parameters:
  - teamId (required): The Team ID
  - email (required): External user's email
  - displayName (optional): Guest's display name
  - sendInvitationMessage (optional): Boolean to send welcome email
Returns: Guest user object
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/team-post-members (for guests)
- Requires Azure AD B2B guest invitation flow

**Note**: This requires specific tenant configuration for external access.

---

## GAPS: Basic Teams Operations (Blocking Issues)

These are the immediate blockers we hit during testing.

### Gap 1: List Joined Teams

**Current State**: Tool `teams.listJoinedTeams` not available
**Workaround Used**: `listMyGroups` (returns M365 Groups, not Teams specifically)
**Problem**: Cannot reliably get Teams that user has joined

**Required Tool**:
```
Tool Name: listJoinedTeams
Description: Get list of Teams the user has joined
Parameters: none (or optional filter)
Returns: Array of Team objects with id, displayName, description
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams
- Endpoint: `GET /me/joinedTeams`

---

### Gap 2: List Team Channels

**Current State**: Tool `teams.listTeamChannels` not available
**Workaround Used**: None - had to skip channel operations entirely
**Problem**: Cannot discover channels within a Team to post messages

**Required Tool**:
```
Tool Name: listTeamChannels
Description: Get list of channels in a Team
Parameters:
  - teamId (required): The Team ID
Returns: Array of Channel objects with id, displayName, description, membershipType
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-list
- Endpoint: `GET /teams/{team-id}/channels`

---

### Gap 3: Send Channel Message (Functional Issues)

**Current State**: Tool `sendChannelMessage` exists but couldn't be used effectively
**Problem**: Without `listJoinedTeams` and `listTeamChannels`, we have no way to obtain valid `teamId` and `channelId` parameters

**Verification Needed**:
- Confirm tool works when valid IDs are provided
- Verify parameter names match Graph API expectations

**Required Parameters** (per Graph API):
```
Tool Name: sendChannelMessage
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - content (required): Message body content
  - contentType (optional): "text" or "html" (default: "text")
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-post-messages
- Endpoint: `POST /teams/{team-id}/channels/{channel-id}/messages`

**ChatMessage Resource**:
- https://learn.microsoft.com/en-us/graph/api/resources/chatmessage

---

### Gap 4: Send Chat Message (Returns Error)

**Current State**: Tool `sendChatMessage` exists but returned error `0` during testing
**Observed Behavior**:
```python
kam_client.send_chat_message(chat_id, message)
# Result: Error "0" - no meaningful error message
```

**Investigation Needed**:
1. Is the `chatId` format correct? (Graph expects specific format)
2. Are permissions sufficient? (Chat.ReadWrite required)
3. Is the request body formatted correctly?

**Required Parameters** (per Graph API):
```
Tool Name: sendChatMessage
Parameters:
  - chatId (required): The chat ID
  - content (required): Message body
  - contentType (optional): "text" or "html"
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/chat-post-messages
- Endpoint: `POST /chats/{chat-id}/messages`

**Request Body Format**:
```json
{
  "body": {
    "contentType": "text",
    "content": "Message content here"
  }
}
```

---

### Gap 5: Get Channel Messages

**Current State**: Unknown if tool exists or works
**Need**: Agents should be able to read channel messages to see project updates

**Required Tool**:
```
Tool Name: getChannelMessages
Description: Get messages from a Teams channel
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - top (optional): Number of messages to return (default: 20)
Returns: Array of ChatMessage objects
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/channel-list-messages
- Endpoint: `GET /teams/{team-id}/channels/{channel-id}/messages`

---

### Gap 6: Reply to Channel Message

**Current State**: Tool `replyToMessage` exists but untested
**Need**: Agents should reply in threads rather than creating new top-level messages

**Required Parameters**:
```
Tool Name: replyToChannelMessage
Parameters:
  - teamId (required): The Team ID
  - channelId (required): The Channel ID
  - messageId (required): Parent message ID to reply to
  - content (required): Reply content
  - contentType (optional): "text" or "html"
```

**Graph API Reference**:
- https://learn.microsoft.com/en-us/graph/api/chatmessage-post-replies
- Endpoint: `POST /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies`

---

## Workflow Gap: Project Kickoff Meeting

### Observation

During the test, the KAM (Anna Kowalski) sent task assignment emails immediately after creating the project. A better workflow would be:

1. Create project
2. **Schedule kickoff meeting with all assigned team members**
3. Post announcement to Teams channel
4. Send detailed task assignments after kickoff

### Required Enhancement

The KAM workflow should schedule an online meeting before sending delegation emails.

**Current Tools Available**:
- `createEvent` - Creates calendar event
- `findMeetingTimes` - Finds available slots

**Missing Integration**:
- Create Teams meeting link as part of event
- Invite all assigned team members automatically
- Post meeting link to project Teams channel

**Suggested Tool Enhancement**:
```
Tool Name: createProjectKickoff
Description: Create a Teams meeting for project kickoff
Parameters:
  - subject (required): Meeting subject
  - projectId (required): Project reference
  - attendees (required): Array of email addresses
  - duration (optional): Meeting duration in minutes (default: 30)
  - body (optional): Meeting agenda/description
Returns: Event object with Teams meeting link
```

This could use existing `createEvent` with `isOnlineMeeting: true` but needs verification that Teams meeting links are generated.

---

## Permission Requirements

For full Teams functionality, the app registration needs these Graph API permissions:

| Permission | Type | Description |
|------------|------|-------------|
| `Team.ReadBasic.All` | Delegated | Read teams user has joined |
| `Channel.ReadBasic.All` | Delegated | Read channel names and descriptions |
| `ChannelMessage.Send` | Delegated | Send messages in channels |
| `ChannelMessage.Read.All` | Delegated | Read channel messages |
| `Chat.ReadWrite` | Delegated | Read and send chat messages |
| `OnlineMeetings.ReadWrite` | Delegated | Create Teams meetings |

---

## Complete Gap Summary

| ID | Tool | Category | Priority | Status |
|----|------|----------|----------|--------|
| A | `createTeamChannel` | Collaboration | **Critical** | Missing |
| B | `addChannelMember` | Collaboration | **Critical** | Missing |
| C | `uploadFileToChannel` | Collaboration | **Critical** | Missing |
| D | `listChannelFiles` | Collaboration | **Critical** | Missing |
| E | `readChannelFile` | Collaboration | **Critical** | Missing |
| F | `inviteGuestToTeam` | Collaboration | Future | Missing |
| 1 | `listJoinedTeams` | Basic | **Critical** | Missing |
| 2 | `listTeamChannels` | Basic | **Critical** | Missing |
| 3 | `sendChannelMessage` | Basic | High | Blocked |
| 4 | `sendChatMessage` | Basic | High | Broken |
| 5 | `getChannelMessages` | Basic | Medium | Unknown |
| 6 | `replyToChannelMessage` | Basic | Medium | Untested |

---

## Testing Checklist

Once gaps are addressed, verify with this test sequence:

### Basic Operations
```
1. [ ] listJoinedTeams returns user's Teams
2. [ ] listTeamChannels returns channels for a Team
3. [ ] sendChannelMessage posts to a channel successfully
4. [ ] getChannelMessages retrieves recent messages
5. [ ] replyToChannelMessage creates threaded reply
6. [ ] sendChatMessage sends 1:1 or group chat message
7. [ ] getChatMessages retrieves chat history
8. [ ] createEvent with isOnlineMeeting creates Teams link
```

### Collaboration Workflow (Full E2E Test)
```
9.  [ ] createTeamChannel creates new project channel
10. [ ] addChannelMember adds user to channel
11. [ ] Added user can see the channel in their Teams client
12. [ ] uploadFileToChannel uploads .md file to Files tab
13. [ ] listChannelFiles shows uploaded files
14. [ ] readChannelFile returns file content
15. [ ] Second NPC can read file uploaded by first NPC
16. [ ] NPCs can chat in channel, see each other's messages
17. [ ] Full workflow: Channel → Add Members → Files → Chat → Meeting → Delivery
```

### Future (Guest Access)
```
16. [ ] inviteGuestToTeam sends invitation to external email
17. [ ] Guest can view channel and files (read-only)
```

---

## References

- **Chat Messages**: https://learn.microsoft.com/en-us/graph/api/chat-post-messages?view=graph-rest-1.0
- **ChatMessage Resource**: https://learn.microsoft.com/en-us/graph/api/resources/chatmessage?view=graph-rest-1.0
- **Channel Messages**: https://learn.microsoft.com/en-us/graph/api/channel-post-messages?view=graph-rest-1.0
- **List Joined Teams**: https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams?view=graph-rest-1.0
- **List Channels**: https://learn.microsoft.com/en-us/graph/api/channel-list?view=graph-rest-1.0
- **Teams Permissions**: https://learn.microsoft.com/en-us/graph/permissions-reference#teams-permissions

---

## Impact on Synthetic Employees

### Current State (Email-Only Fallback)
```
Agent A sends email → Agent B reads email → Agent B sends email → ...
```
- No shared workspace
- No visibility into what others are doing
- Documents sent as attachments, no single source of truth
- No real-time collaboration
- Client completely out of the loop

### Desired State (True NPC Collaboration)
```
┌─────────────────────────────────────────────────────────────────┐
│  Project Channel: "Prysmian-41e06072"                           │
│                                                                  │
│  [Kickoff Meeting Scheduled - All Invited]                       │
│                                                                  │
│  Files:                                                          │
│  📄 project-brief.md      (KAM)                                  │
│  📄 research-findings.md  (Researcher)                           │
│  📄 draft-v1.md           (Writer)                               │
│  📄 final-approved.md     (Editor)                               │
│                                                                  │
│  Chat:                                                           │
│  [Anna] Project started! See brief in Files                      │
│  [Christina] @Bruno I uploaded draft-v1.md for review            │
│  [Bruno] Looks good! Minor fixes in line 42                      │
│  [Christina] Fixed and updated the file                          │
│  [François] Approved! ✓                                          │
│  [Anna] Delivering to client now                                 │
│                                                                  │
│  [Future: Client "Marcus" can see progress in real-time]         │
└─────────────────────────────────────────────────────────────────┘
```

### What This Enables

| Capability | Impact |
|------------|--------|
| **Dedicated project channel** | Clean separation per project, no noise |
| **Shared Files tab** | Single source of truth, version history |
| **.md research files** | Researchers share findings, others consume |
| **Real-time chat** | NPCs discuss, @mention, react instantly |
| **Threaded replies** | Organized conversations per topic |
| **Kickoff meetings** | Proper project initiation with all members |
| **Client visibility** | (Future) Transparency, trust, fewer emails |

### The NPC Behavior We Want

```python
# When KAM receives external project request:

1. kam.create_channel(f"Project-{project_id}-{client_name}")

2. # ADD TEAM MEMBERS TO CHANNEL (critical step!)
   kam.add_channel_member(channel, "christina.hall@...")   # Marketing Writer
   kam.add_channel_member(channel, "bruno.dupont@...")     # Proofreader
   kam.add_channel_member(channel, "francois.moreau@...")  # Editor
   # Now they can all see and access the channel

3. kam.upload_file(channel, "project-brief.md", requirements)
4. kam.schedule_meeting(channel, "Project Kickoff", team_members)
5. kam.post_message(channel, "Welcome team! See brief in Files. Kickoff at 2pm")
6. kam.assign_tasks(team_members)  # Still via email for formal record

# Team members then (they can now see the channel!):
7.  researcher.read_file(channel, "project-brief.md")
8.  researcher.upload_file(channel, "research-findings.md", findings)
9.  researcher.post_message(channel, "@Christina research uploaded!")
10. writer.read_file(channel, "research-findings.md")
11. writer.upload_file(channel, "draft-v1.md", content)
12. writer.post_message(channel, "@Bruno draft ready for review")
# ... and so on, visible collaboration
```

This is how NPCs should work - like real teammates in a shared workspace.

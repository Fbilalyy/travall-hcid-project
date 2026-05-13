# IE48L HCID - Assignment 2: Hierarchical Task Analysis

## Task 1: HTA Based on a User's Mental Model

I asked my friend to show me how they would share a photo via email on their iPhone. They picked up the phone and went straight to the Photos app, which made sense since they already had a specific photo in mind. What I found interesting was that they didn't hesitate at all with the first few steps - opening the photo and hitting the share button seemed completely automatic to them. But once the share sheet appeared, they paused for a second to scan through the options before tapping on Mail.

After selecting Mail, a new email compose window opened with the photo already attached at the bottom. My friend then typed in the recipient's email address, skipped the subject line entirely, and hit Send right away. When I asked about the subject line, they shrugged and said "I never bother with that for photos." The whole thing took about 20 seconds once they found the photo they wanted.

Below is the textual representation of the HTA based on what I observed, followed by the diagram.

### Textual HTA - Method 1 (Photo Gallery Approach)

**0. Share a photo via email on iPhone**
Plan 0: Do 1, then 2, then 3, then 4 in order

**1. Locate and open the photo**
Plan 1: Do 1.1, then 1.2, then 1.3

- 1.1 Unlock iPhone
- 1.2 Open the Photos app
- 1.3 Browse and tap on the desired photo to view it full-screen

**2. Initiate the sharing process**
Plan 2: Do 2.1, then 2.2

- 2.1 Tap the Share button (square icon with an upward arrow at the bottom-left)
- 2.2 Select "Mail" from the share sheet options

**3. Compose the email**
Plan 3: Do 3.1. If needed, do 3.2 and/or 3.3

- 3.1 Type the recipient's email address in the "To" field
- 3.2 Enter a subject line (optional)
- 3.3 Write a message in the body (optional)

**4. Send the email**

- 4.1 Tap the blue Send button (top-right corner)

### HTA Diagram - Method 1

```
                    ┌─────────────────────────┐
                    │  0. Share a photo via    │
                    │  email on iPhone         │
                    └───────────┬─────────────┘
                                │
              Plan 0: Do 1, then 2, then 3, then 4
                                │
         ┌──────────┬───────────┼───────────┬──────────┐
         │          │           │           │          │
   ┌─────┴─────┐ ┌──┴──────┐ ┌──┴────────┐ ┌┴─────────┐
   │ 1. Locate │ │2. Start │ │3. Compose │ │4. Send   │
   │ & open    │ │sharing  │ │the email  │ │the email │
   │ the photo │ │process  │ │           │ │          │
   └─────┬─────┘ └──┬──────┘ └──┬────────┘ └┬─────────┘
         │          │           │           │
  Plan 1:│   Plan 2:│    Plan 3:│           │
  1.1,   │   2.1,   │    Do 3.1.│       ┌───┴───┐
  1.2,   │   then   │  Opt. 3.2 │       │ 4.1   │
  1.3    │   2.2    │  and/or   │       │ Tap   │
         │          │  3.3      │       │ Send  │
   ┌──┬──┴┐   ┌──┬──┘     ┌──┬──┘       └───────┘
   │  │   │   │  │        │  │  │
 ┌─┴┐┌┴┐┌─┴┐┌─┴┐┌┴──┐ ┌──┴┐┌┴┐┌┴──┐
 │  ││  ││  ││  ││   │ │   ││ ││   │
 │1.││1.││1.││2.││2.2│ │3.1││3││3.3│
 │1 ││2 ││3 ││1 ││Sel│ │To ││.││Msg│
 │Un││Op││Br││Ta││ect│ │   ││2││   │
 │lo││en││ow││p ││Mai│ │   ││ ││   │
 │ck││Ph││se││Sh││l  │ │   ││S││   │
 │  ││ot││& ││ar││   │ │   ││u││   │
 │  ││os││ta││e ││   │ │   ││b││   │
 │  ││  ││p ││  ││   │ │   ││j││   │
 └──┘└──┘└──┘└──┘└───┘ └───┘└─┘└───┘
```

---

## Task 2: Alternative HTA - Mail App Approach

For the second method, I analyzed the process of sharing a photo by starting from the Mail app instead. This is what I would consider the "email-first" approach: you open Mail, compose a new message, and then attach a photo from the gallery. I tried this on my own iPhone to map out all the steps carefully.

One thing that stands out right away is that this method has more steps in the attachment phase. After you start composing the email and fill in the recipient, you need to figure out how to add a photo. On iOS, you can either tap the attachment icon in the toolbar above the keyboard, or you can tap inside the email body and sometimes get an "Insert Photo" option. Once you get into the photo picker, you have to navigate your photo library to find the right image - and if you have thousands of photos, this part can be genuinely frustrating. There's no guarantee the photo you want is immediately visible, so you might need to scroll through albums or use the search function.

I noticed this method feels a bit more cumbersome than Method 1. In Method 1, you start with the photo already in front of you, so there's no searching involved. Here, the searching happens after you've already set up the email, which can feel like going backwards.

### Textual HTA - Method 2 (Mail App Approach)

**0. Share a photo via email on iPhone**
Plan 0: Do 1, then 2, then 3, then 4 in order

**1. Open Mail and start composing**
Plan 1: Do 1.1, then 1.2

- 1.1 Unlock iPhone and open the Mail app
- 1.2 Tap the Compose button (pencil/paper icon at bottom-right)

**2. Fill in email details**
Plan 2: Do 2.1. If needed, do 2.2 and/or 2.3

- 2.1 Type the recipient's email address in the "To" field
- 2.2 Enter a subject line (optional)
- 2.3 Write a message in the body (optional)

**3. Attach the photo**
Plan 3: Do 3.1, then 3.2, then 3.3. If photo is not immediately visible, also do 3.3.1

- 3.1 Tap the attachment icon in the toolbar (or tap the "<" arrow to expand toolbar, then tap the photo/camera icon)
- 3.2 Select "Photo Library" from the options
- 3.3 Find and select the desired photo
    - Plan 3.3: Do 3.3.1 if the photo is not in the recent view. Then do 3.3.2
    - 3.3.1 Navigate through albums or use the search bar to locate the photo
    - 3.3.2 Tap on the photo to attach it

**4. Send the email**

- 4.1 Review the composed email with the attached photo
- 4.2 Tap the blue Send button

### HTA Diagram - Method 2

```
                    ┌─────────────────────────┐
                    │  0. Share a photo via    │
                    │  email on iPhone         │
                    └───────────┬─────────────┘
                                │
              Plan 0: Do 1, then 2, then 3, then 4
                                │
         ┌──────────┬───────────┼───────────┬──────────┐
         │          │           │           │          │
   ┌─────┴─────┐ ┌──┴────────┐ ┌┴──────────┐┌┴─────────┐
   │ 1. Open   │ │2. Fill in │ │3. Attach  ││4. Send   │
   │ Mail &    │ │email      │ │the photo  ││the email │
   │ compose   │ │details    │ │           ││          │
   └─────┬─────┘ └──┬────────┘ └┬──────────┘└┬─────────┘
         │          │           │            │
  Plan 1:│   Plan 2:│    Plan 3:│        ┌───┼───┐
  1.1,   │   Do 2.1.│    3.1,   │        │       │
  then   │  Opt.2.2 │    3.2,   │     ┌──┴──┐ ┌──┴──┐
  1.2    │  and/or  │    3.3    │     │ 4.1 │ │ 4.2 │
         │  2.3     │           │     │Revi-│ │ Tap │
   ┌──┬──┘    ┌──┬──┘     ┌──┬──┘     │ew   │ │Send │
   │  │       │  │  │     │  │  │     └─────┘ └─────┘
 ┌─┴┐┌┴──┐ ┌─┴┐┌┴┐┌┴──┐┌─┴┐┌┴──┐┌┴──┐
 │1.││1.2│ │2.││2││2.3││3.││3.2││3.3│
 │1 ││Com│ │1 ││.││Msg││1 ││Pho││Fin│
 │Op││pos│ │To││2││   ││At││to ││d &│
 │en││e  │ │  ││S││   ││ta││Lib││sel│
 │Ma││   │ │  ││u││   ││ch││rar││ect│
 │il││   │ │  ││b││   ││  ││y  ││   │
 └──┘└───┘ └──┘└─┘└───┘└──┘└───┘└─┬─┘
                                    │
                          Plan 3.3: If not visible,
                          do 3.3.1. Then do 3.3.2
                                    │
                              ┌─────┼─────┐
                              │           │
                          ┌───┴───┐  ┌────┴────┐
                          │ 3.3.1 │  │  3.3.2  │
                          │Browse │  │  Tap to  │
                          │albums/│  │  attach  │
                          │search │  │         │
                          └───────┘  └─────────┘
```

---

## Task 3: Comparison of the Two HTA Diagrams

### Structural Similarities

Both methods share the same top-level goal: sharing a photo via email on an iPhone. They also share certain subtasks that are essentially identical regardless of which approach is taken. Composing the email (filling in the "To" field, optionally adding a subject and body text) and the final act of tapping Send appear in both hierarchies. In a sense, the core email-related operations are constant - what changes is the order and method of getting the photo into the email.

Both HTAs also begin with unlocking the phone, which is a trivial but necessary first step. And in both cases, the user ends up at the same final screen: a composed email with an attached photo, ready to be sent.

### Key Differences

The most significant difference lies in where the photo selection happens within the task flow. In Method 1, the user starts with the photo. They find it first, open it, and then the sharing mechanism (the share sheet) handles the email setup. The photo is, in effect, already "attached" before the user even thinks about the email. In Method 2, the user starts with the email and has to go back to find the photo midway through the process - it reverses the order of the two main operations.

This difference in ordering leads to a deeper subtask structure in Method 2's attachment phase. Method 1 avoids the need for a dedicated "find the photo in your library" step because the user already selected it at the beginning. Method 2 introduces subtasks 3.3.1 (browsing albums or searching) and 3.3.2 (tapping to attach), creating an additional level in the hierarchy.

In terms of HTA structure, Method 1 is flatter with fewer levels, while Method 2 has a deeper hierarchy under the "Attach the photo" branch. As Lane, Stanton, and Harrison (2006) point out, deeper task hierarchies with more subtask levels tend to introduce more opportunities for errors at each decomposition level.

### Execution Speed

Method 1 is generally faster. The user goes directly from the photo to the email, and the attachment is handled automatically by iOS when the share sheet opens the Mail compose window. There are roughly 7 leaf-level operations in Method 1 compared to about 9-10 in Method 2. The extra steps in Method 2 - particularly finding the attachment icon, selecting Photo Library, and then navigating to the correct photo - add both time and cognitive load.

That said, if the user already has the Mail app open and happens to know exactly where the photo is in their library, the speed difference becomes less significant. Context matters here.

### Vulnerability to Mistakes

Method 2 is more prone to errors, mainly due to the photo-finding step. When a user starts from the Mail app, they need to recall which photo they want to send and then locate it from a potentially large library. This is a retrieval task that can fail in several ways: the user might select the wrong photo, have difficulty finding the attachment icon (which is not always immediately visible on iOS), or accidentally close the photo picker before completing the selection.

Method 1 reduces this vulnerability because the user starts by looking at the actual photo. There is almost no chance of sending the wrong image since it is right in front of them throughout the process. The share sheet is also a well-known iOS pattern that most users recognize instinctively - the square-with-arrow icon is consistent across almost every app on the platform.

However, Method 1 does have a subtle weakness: if the user taps Mail on the share sheet and then realizes they wanted to make changes to the email draft (e.g., go back and select a different photo or add more photos), it can be awkward to navigate back. Method 2, on the other hand, keeps the user in the email compose view throughout, making it easier to manage multiple attachments or edit the email before attaching.

### Simplicity

Method 1 is simpler from the user's perspective. It requires fewer decisions and the task flows in a more natural direction: "I have this photo, I want to share it." The system (iOS share sheet) handles the integration between the photo and the email app. The cognitive load is lower because each step follows logically from the previous one without requiring the user to switch contexts.

Method 2 requires the user to hold the intention of "I need to attach a photo" in working memory while they navigate away from the compose screen into the photo library. This context switch can be disorienting, especially for users who are less comfortable with technology. An older user or someone with cognitive difficulties might lose track of what they were doing during the photo selection step.

In summary, Method 1 (Photo Gallery approach) is faster, simpler, and less error-prone for sharing a single photo. Method 2 (Mail App approach) offers more flexibility for composing a detailed email with multiple attachments but comes at the cost of added complexity and a higher chance of mistakes during the photo selection phase. For the specific task of "sharing a photo via email," Method 1 aligns better with how most people naturally think about the task - starting with the content they want to share rather than the tool they will use to share it.

### References

Lane, R., Stanton, N. A., & Harrison, D. (2006). Applying hierarchical task analysis to medication administration errors. *Applied Ergonomics*, 37(5), 669-679.

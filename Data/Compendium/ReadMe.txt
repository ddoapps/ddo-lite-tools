*** IMPORTANT ***

If you need to rename a quest or challenge:

Both the Quest and Challenge data structures contain an ID field, which is used for quest
and challenge lookups and is required to retain character progress.

At launch, all IDs are exactly the same as the quest/challenge name, so no IDs appear in
the original versions of Quests.txt and Challenges.txt.

If, for whatever reason, a quest or challenge name needs to be changed, you must add a line 
to set an ID field to the original name. Without this, existing compendium files won't be 
able to connect character progress with those quests.

For example, let's say we wanted to remove the "!" from the Ghola Fan quest name. Originally
the quest appears like this:

QuestName: Bring Me the Head of Ghola-Fan!
Pack: Restless Isles
Patron: The Free Agents
Favor: 6
Level: 10

It's fine to change it, but then we need to add an ID line to preserve the original value:

QuestName: Bring Me the Head of Ghola-Fan
ID: Bring Me the Head of Ghola-Fan!
Pack: Restless Isles
Patron: The Free Agents
Favor: 6
Level: 10

If we then change it again, the ID field stays the same. The ID must always match the
ORIGINAL value for the quest or challenge name.

Note that IDs are case-sensitive, so even if you only change the capitalization of a name,
you'll have to create an ID field.

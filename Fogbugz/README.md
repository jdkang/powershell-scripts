This was quick script to check FogBugz for unassigned items of a certain priority and notify HipChat.

The FogBugz API scaffolding isn't fully flushe dout, but the basic cmd wrappers should be good enough to do most things in conjunction with the [Fogbugz API documentation](http://help.fogcreek.com/8202/xml-api)

The script can map different projects to different rooms in the case you have multiple product teams that may rely on a single queue (milestone) of work.

The FB query is hard-coded so you will to adjust it as needed.
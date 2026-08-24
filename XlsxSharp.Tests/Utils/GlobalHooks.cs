// Replaces NUnit's assembly-level [Culture("en-US")]. TUnit's [Culture] attribute (unlike a
// plain event-receiver-based reset) properly scopes and restores culture per test, so a
// method-level [Culture("xx-YY")] override correctly reverts to this default afterwards instead
// of leaking into whichever test runs next.
[assembly: Culture("en-US")]

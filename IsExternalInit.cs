// net48 does not ship this type; C# 9 init-only setters require it.
// The NATS.Net option objects (NatsOpts, NatsAuthOpts, NatsTlsOpts) use
// init-only setters, so the add-in must compile as C# 9.
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}

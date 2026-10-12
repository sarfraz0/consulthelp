# sendshards implementation notes

[Documentation index](README.md)

These notes preserve function help and explanatory comments removed from the script. Original source line numbers refer to the pre-extraction file. All executable tokens and string literals remain unchanged.

## ConvertTo-NativeArgument

Original line 281: Windows CRT quoting: double slashes before a quote and before the closing quote.

## Stop-OwnedProcess

Original line 330: Never kill by image name. The PID belongs to the Process object started here.

## Send-ItemCommit

Original line 1473: Exact source OID, one branch refspec, no force, no tag propagation, hooks active.

## Invoke-TransferPush

Original line 1510: Save the original start before contacting the remote or making Git mutations.

## Invoke-PlanPublication

Original line 1617: Revalidate the frozen view after acquiring the tool lock.

## Entry point

Original line 1633: Thin executable boundary: no dot-sourcing interface, success objects, or custom prompts.

Original line 1650: JSON reader assembly is loaded here before reading the input document.

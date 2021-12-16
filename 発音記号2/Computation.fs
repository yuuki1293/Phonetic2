module Phonetic.Computation

let (>>=) x f =
    match x with
    | Some x -> f x
    | None -> None

let ret x = Some x

type MaybeBuilder() =
    member _.Bind(x, f) = x >>= f
    member _.Return x = Some x
    member _.Zero _ = None
    member _.ReturnFrom x = x 

let maybe = MaybeBuilder()

from spacy.matcher import Matcher


def IsImperative(text, doc, nlp):
    i = 1
    if doc[1].pos_ == "PUNCT":
        i = 2

    # If the sentence not is question
    # First word is VB or MD, e.g. "Give it to me."
    if doc[0].pos_ == "VERB":
        return True, 1
    # E.g. "Dave, stop.", "Man, stop!"
    elif (
        doc[0].pos_ == "PRON" or doc[0].pos_ == "PROPN" or doc[0].tag_ == "NN"
    ) and doc[i].pos_ == "VERB":
        return True, 2

    return False, 0


def GetVerbChunk(nlp, doc):
    matcher = Matcher(nlp.vocab)

    pattern = [{"TAG": "VB"}, {"TAG": "RP"}]
    matcher.add("VB+RP", [pattern])
    matches = matcher(doc)

    if len(matches) == 1:
        return doc[matches[0][1] : matches[0][2]]
    else:
        return None


def MatcherF(nlp, doc):
    matcher = Matcher(nlp.vocab)

    # Add pattern to matcher
    pattern_1 = [{"POS": "PRON"}, {"POS": "VERB"}]
    pattern_2 = [{"POS": "PROPN"}, {"POS": "VERB"}]
    pattern_3 = [{"TAG": "NN"}, {"POS": "VERB"}]

    matcher.add("PRON+VERB", [pattern_1])
    matcher.add("PROPN+VERB", [pattern_2])
    matcher.add("NN+VERB", [pattern_3])
    # call the matcher to find matches
    matches = matcher(doc)

    if len(matches) == 1:
        return doc[matches[0][1] : matches[0][2]]
    else:
        return None


def GetActionAndObject(nlp, command):
    doc = nlp(command)

    is_imperative, n_rule = IsImperative(command, doc, nlp)

    if n_rule == 1:
        verb_chunk = GetVerbChunk(nlp, doc)

        if verb_chunk is not None:
            if len(list(doc.noun_chunks)) > 0 and len(verb_chunk) == 1:
                return True, verb_chunk, str(list(doc.noun_chunks)[0]), doc
            elif len(list(doc.noun_chunks)) == 0:
                if doc[1].pos_ == "PUNCT":
                    return True, verb_chunk, str(doc[2]), doc
                else:
                    if len(doc) - 1 == len(verb_chunk):
                        return True, verb_chunk, "you", doc
                    else:
                        return True, verb_chunk, str(doc[1]), doc
        else:
            if len(list(doc.noun_chunks)) > 0:
                return True, str(doc[0]), str(list(doc.noun_chunks)[0]), doc
            else:
                return True, str(doc[0]), "you", doc

    elif n_rule == 2:
        if doc[1].pos_ == "VERB":
            return True, str(doc[1]), str(doc[0]), doc
        else:
            return True, str(doc[2]), str(doc[0]), doc
    else:
        return False, None, None, doc

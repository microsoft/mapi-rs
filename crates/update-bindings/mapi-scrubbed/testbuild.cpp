// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "MAPIConstants.h"

#include <cstdint>

static_assert(MAPI_E_NOT_FOUND == ((HRESULT)0x8004010F));
static_assert(PT_MV_BINARY == 0x1102);
static_assert(PR_CONVERSATION_TOPIC_A == (0x0070001E));
static_assert(BOOKMARK_BEGINNING == 0);
static_assert(TABLE_SORT_CATEG_MAX == 4);
static_assert(static_cast<std::int32_t>(PRIO_NONURGENT) == -1);

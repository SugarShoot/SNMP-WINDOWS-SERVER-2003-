//

//  dfsrperf.h
//
//  Offset definition file for exensible counter objects and counters
//
//  These "relative" offsets must start at 0 and be multiples of 2 (i.e.
//  even numbers). In the Open Procedure, they will be added to the
//  "First Counter" and "First Help" values fo the device they belong to,
//  in order to determine the  absolute location of the counter and
//  object names and corresponding help text in the registry.
//
//  This file is used by the extensible counter DLL code as well as the
//  counter name and help text definition file (.INI) file that is used
//  by LODCTR to load the names into the registry.
//

#ifndef DFSRPERF_H
#define DFSRPERF_H

#define PERFOBJ_VOLUMES                        0

#define PERFOBJ_VOLUME_USN_RECORDS_EXAMINED    2
#define PERFOBJ_VOLUME_USN_RECORDS_ACCEPTED    4
#define PERFOBJ_VOLUME_JOURNAL_UNREAD_PCT      6
#define PERFOBJ_VOLUME_DB_COMMITS              8
#define PERFOBJ_VOLUME_DB_LOOKUPS              10

#define PERFOBJ_CONTENT                        12

#define PERFOBJ_CONTENT_STAGING_FILES_GEN      14
#define PERFOBJ_CONTENT_STAGING_BYTES_GEN      16
#define PERFOBJ_CONTENT_STAGING_FILES_DEL      18
#define PERFOBJ_CONTENT_STAGING_BYTES_DEL      20
#define PERFOBJ_CONTENT_STAGING_SPACE_INUSE    22

#define PERFOBJ_CONTENT_CONFLICT_FILES_GEN     24
#define PERFOBJ_CONTENT_CONFLICT_BYTES_GEN     26
#define PERFOBJ_CONTENT_CONFLICT_FILES_DEL     28
#define PERFOBJ_CONTENT_CONFLICT_BYTES_DEL     30
#define PERFOBJ_CONTENT_CONFLICT_SPACE_INUSE   32
#define PERFOBJ_CONTENT_CONFLICT_FOLDER_CLEANUP 34

#define PERFOBJ_CONTENT_FILE_INSTALLS_SUCCESS  36
#define PERFOBJ_CONTENT_FILE_INSTALLS_RETRY    38
#define PERFOBJ_CONTENT_UPDATES_DROPPED        40

#define PERFOBJ_CONTENT_DELETED_FILES_GEN      42
#define PERFOBJ_CONTENT_DELETED_BYTES_GEN      44
#define PERFOBJ_CONTENT_DELETED_FILES_DEL      46
#define PERFOBJ_CONTENT_DELETED_BYTES_DEL      48
#define PERFOBJ_CONTENT_DELETED_SPACE_IN_USE   50

#define PERFOBJ_CONTENT_INBOUND_TOTAL_FILE_TRANSFERS           52
#define PERFOBJ_CONTENT_INBOUND_TOTAL_UNCOMPRESS_FILE_SIZE     54
#define PERFOBJ_CONTENT_INBOUND_TOTAL_BYTES                    56
#define PERFOBJ_CONTENT_RDC_INBOUND_TOTAL_FILE_TRANSFERS       58
#define PERFOBJ_CONTENT_RDC_INBOUND_TOTAL_UNCOMPRESS_FILE_SIZE 60
#define PERFOBJ_CONTENT_RDC_INBOUND_TOTAL_COMPRESS_FILE_SIZE   62
#define PERFOBJ_CONTENT_RDC_INBOUND_TOTAL_BYTES                64
#define PERFOBJ_CONTENT_INBOUND_RELATIVE_SAVINGS               66

#define PERFOBJ_CONNECTIONS                    68

#define PERFOBJ_CONNECTION_BYTES_RECEIVED      70
#define PERFOBJ_CONNECTION_INBOUND_TOTAL_FILE_TRANSFERS           72
#define PERFOBJ_CONNECTION_INBOUND_TOTAL_UNCOMPRESS_FILE_SIZE     74
#define PERFOBJ_CONNECTION_INBOUND_TOTAL_BYTES                    76
#define PERFOBJ_CONNECTION_AVG_BYTES_RECEIVED_PER_SECOND          78
#define PERFOBJ_CONNECTION_RDC_INBOUND_TOTAL_FILE_TRANSFERS       80
#define PERFOBJ_CONNECTION_RDC_INBOUND_TOTAL_UNCOMPRESS_FILE_SIZE 82
#define PERFOBJ_CONNECTION_RDC_INBOUND_TOTAL_COMPRESS_FILE_SIZE   84
#define PERFOBJ_CONNECTION_RDC_INBOUND_TOTAL_BYTES                86
#define PERFOBJ_CONNECTION_INBOUND_RELATIVE_SAVINGS               88

#endif

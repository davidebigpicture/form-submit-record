--------------------------------------------------------------------------
-- form_submit_ledger.sql
-- Durable form-submission ledger schema (additive; no existing tables are
-- altered). Backs cFormSubmitLedger.inc and the admin dashboard columns.
--
-- See docs/durable_submission_schema_plan.md for rationale and operation
-- semantics. Run once per client schema. Safe to re-run only after the
-- objects are dropped; there are no CREATE OR REPLACE statements for tables.
--------------------------------------------------------------------------

-- =======================================================================
-- Sequences
-- =======================================================================
create sequence seq_form_submit_operation_id      start with 1 increment by 1 nocache;
create sequence seq_form_submit_operation_item_id  start with 1 increment by 1 nocache;
create sequence seq_form_submit_event_id           start with 1 increment by 1 nocache;

-- =======================================================================
-- form_submit_record : one durable row per submission (keyed to app_renew)
-- =======================================================================
create table form_submit_record (
    app_renew_id             number(12)            not null,
    app_renew_type           varchar2(50 char)     not null,
    membership_id            number(12),
    durable_state            varchar2(30 char)     default 'draft' not null,
    legacy_app_renew_status  varchar2(30 char),
    ready_to_process         char(1 char)          default 'N' not null,
    submitted_at             timestamp(6) with local time zone,
    completed_at             timestamp(6) with local time zone,
    last_error              varchar2(2000 char),
    retry_after              timestamp(6) with local time zone,
    created_at               timestamp(6) with local time zone default systimestamp not null,
    updated_at               timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_record primary key (app_renew_id),
    constraint ck_fsr_ready check (ready_to_process in ('Y','N')),
    constraint ck_fsr_state check (durable_state in (
        'draft','processing','awaiting_payment','payment_complete',
        'completed','failed','abandoned'))
);

comment on table  form_submit_record is 'Durable per-submission state; mirrors but does not replace app_renew.';
comment on column form_submit_record.durable_state is 'Execution lifecycle state, independent of legacy app_renew_status.';
comment on column form_submit_record.ready_to_process is 'Y when a recovery worker should (re)attempt outstanding operations.';

-- =======================================================================
-- form_submit_operation : one row per operation per submission
-- operation_key in (payment, files, records, pdf, emails, billing, custom_code)
-- =======================================================================
create table form_submit_operation (
    operation_id        number(12)            not null,
    app_renew_id        number(12)            not null,
    operation_key       varchar2(50 char)     not null,
    configured          char(1 char)          default 'Y' not null,
    source_scope        varchar2(30 char)     default 'form_config' not null,
    source_label        varchar2(200 char),
    trigger_field       varchar2(100 char),
    trigger_operator    varchar2(10 char),
    trigger_value       varchar2(400 char),
    state               varchar2(30 char)     default 'ready' not null,
    expected_count      number(6)             default 0 not null,
    done_count          number(6)             default 0 not null,
    failed_count        number(6)             default 0 not null,
    attempts            number(6)             default 0 not null,
    notes               varchar2(2000 char),
    last_error          varchar2(2000 char),
    started_at          timestamp(6) with local time zone,
    completed_at        timestamp(6) with local time zone,
    created_at          timestamp(6) with local time zone default systimestamp not null,
    updated_at          timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_operation primary key (operation_id),
    constraint uq_form_submit_operation unique (app_renew_id, operation_key),
    constraint ck_fso_configured check (configured in ('Y','N')),
    constraint ck_fso_scope check (source_scope in ('form_config','triggered','not_configured')),
    constraint ck_fso_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'))
);

comment on column form_submit_operation.source_scope is 'form_config = form-wide default; triggered = field-conditional; not_configured = NA.';
comment on column form_submit_operation.state is 'Mockup operation vocabulary; see form_submit_tool.html OP_STATE.';

-- =======================================================================
-- form_submit_operation_item : item-level detail (each file / each email)
-- =======================================================================
create table form_submit_operation_item (
    item_id             number(12)            not null,
    operation_id        number(12),
    app_renew_id        number(12)            not null,
    operation_key       varchar2(50 char)     not null,
    item_key            varchar2(400 char)    not null,
    item_label          varchar2(400 char),
    source_scope        varchar2(30 char)     default 'form_config' not null,
    trigger_field       varchar2(100 char),
    trigger_operator    varchar2(10 char),
    trigger_value       varchar2(400 char),
    state               varchar2(30 char)     default 'ready' not null,
    temp_location       varchar2(1000 char),
    expected_location   varchar2(1000 char),
    permanent_location  varchar2(1000 char),
    content_hash        varchar2(128 char),
    size_bytes          number(18),
    attempts            number(6)             default 0 not null,
    notes               varchar2(2000 char),
    last_error          varchar2(2000 char),
    completed_at        timestamp(6) with local time zone,
    created_at          timestamp(6) with local time zone default systimestamp not null,
    updated_at          timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_operation_item primary key (item_id),
    constraint uq_form_submit_operation_item unique (app_renew_id, operation_key, item_key),
    constraint ck_fsi_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'))
);

-- =======================================================================
-- form_submit_event : append-only troubleshooting timeline
-- =======================================================================
create table form_submit_event (
    event_id        number(12)            not null,
    app_renew_id    number(12)            not null,
    operation_key   varchar2(50 char),
    item_id         number(12),
    event_type      varchar2(100 char)    not null,
    severity        varchar2(20 char)     default 'info' not null,
    message         varchar2(4000 char)   not null,
    payload_json    clob,
    created_at      timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_event primary key (event_id),
    constraint ck_fse_severity check (severity in ('debug','info','warn','error'))
);

-- =======================================================================
-- form_submit_config_snapshot : immutable config captured at submit time
-- =======================================================================
create table form_submit_config_snapshot (
    app_renew_id    number(12)            not null,
    app_renew_type  varchar2(50 char)     not null,
    config_version  varchar2(100 char),
    config_json     clob                  not null,
    created_at      timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_config_snapshot primary key (app_renew_id)
);

-- =======================================================================
-- Indexes for admin list + recovery worker scans
-- =======================================================================
create index ix_fsr_state        on form_submit_record   (durable_state, updated_at);
create index ix_fsr_ready        on form_submit_record   (ready_to_process, retry_after);
create index ix_fso_app          on form_submit_operation (app_renew_id);
create index ix_fso_key_state    on form_submit_operation (operation_key, state);
create index ix_fsi_app          on form_submit_operation_item (app_renew_id, operation_key);
create index ix_fse_app          on form_submit_event    (app_renew_id, created_at);

-- =======================================================================
-- Foreign keys (no cascade: keep ledger rows even if app_renew is purged,
-- and never let a child-table problem block a submission write).
-- Operations/items/events intentionally reference form_submit_record.
-- =======================================================================
alter table form_submit_operation
    add constraint fk_fso_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

alter table form_submit_operation_item
    add constraint fk_fsi_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

alter table form_submit_event
    add constraint fk_fse_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

-- =======================================================================
-- v_form_submit_admin : flattened view for apprenewadmin.asp and dashboards.
-- One row per submission with the operation states the admin grid surfaces.
-- =======================================================================
create or replace view v_form_submit_admin as
select
    r.app_renew_id,
    r.app_renew_type,
    r.membership_id,
    r.durable_state,
    r.legacy_app_renew_status,
    r.ready_to_process,
    r.submitted_at,
    r.completed_at,
    r.updated_at,
    (select o.state from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'payment')      payment_state,
    (select o.state from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'files')         files_state,
    (select o.done_count || '/' || o.expected_count from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'files')         files_count,
    (select o.state from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'pdf')           pdf_state,
    (select o.state from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'emails')        emails_state,
    (select o.done_count || '/' || o.expected_count from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'emails')        emails_count,
    (select o.state from form_submit_operation o
       where o.app_renew_id = r.app_renew_id and o.operation_key = 'records')       records_state,
    (select count(*) from form_submit_event e
       where e.app_renew_id = r.app_renew_id
         and e.event_type = 'awaiting_updates_released')                            released_count
from form_submit_record r;

-- =======================================================================
-- Rollback (manual):
--   drop view v_form_submit_admin;
--   drop table form_submit_config_snapshot purge;
--   drop table form_submit_event purge;
--   drop table form_submit_operation_item purge;
--   drop table form_submit_operation purge;
--   drop table form_submit_record purge;
--   drop sequence seq_form_submit_event_id;
--   drop sequence seq_form_submit_operation_item_id;
--   drop sequence seq_form_submit_operation_id;
-- =======================================================================

--------------------------------------------------------------------------
-- form_submit_ledger.sql
-- Durable form-submission ledger schema (additive; no existing tables are
-- altered). Backs cFormSubmitLedger.inc and the admin dashboard columns.
--
-- Canonical operation states match form_submit_tool.html OP_STATE and
-- docs/durable_submission_schema_plan.md. Run once per client schema.
--------------------------------------------------------------------------

-- =======================================================================
-- Sequences
-- =======================================================================
create sequence seq_form_submit_operation_id       start with 1 increment by 1 nocache;
create sequence seq_form_submit_operation_item_id  start with 1 increment by 1 nocache;
create sequence seq_form_submit_event_id           start with 1 increment by 1 nocache;

-- =======================================================================
-- form_submit_record : one durable row per submission (keyed to app_renew)
-- =======================================================================
create table form_submit_record (
    app_renew_id             number(12)            not null,
    app_renew_type           varchar2(50 char)     not null,
    membership_id            number(12),
    durable_state            varchar2(30 char)     not null,
    legacy_app_renew_status  varchar2(30 char),
    submitted_at             timestamp(6) with local time zone,
    ready_to_process         char(1 char)          default 'N' not null,
    processing_started_at    timestamp(6) with local time zone,
    completed_at             timestamp(6) with local time zone,
    abandoned_at             timestamp(6) with local time zone,
    current_operation_key    varchar2(50 char),
    last_error               varchar2(4000 char),
    retry_after              timestamp(6) with local time zone,
    created_at               timestamp(6) with local time zone default systimestamp not null,
    updated_at               timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_record primary key (app_renew_id),
    constraint ck_fsr_ready_to_process check (ready_to_process in ('Y','N')),
    constraint ck_fsr_durable_state check (durable_state in (
        'draft','awaiting_payment','payment_complete','processing',
        'completed','failed','abandoned'
    ))
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
    configured          char(1 char)          not null,
    source_scope        varchar2(30 char)     not null,
    source_label        varchar2(200 char),
    trigger_field       varchar2(100 char),
    trigger_operator    varchar2(20 char),
    trigger_value       varchar2(400 char),
    state               varchar2(30 char)     not null,
    expected_count      number(9)             default 0 not null,
    done_count          number(9)             default 0 not null,
    failed_count        number(9)             default 0 not null,
    attempts            number(5)             default 0 not null,
    max_attempts        number(5)             default 0 not null,
    next_attempt_at     timestamp(6) with local time zone,
    job_key             varchar2(100 char),
    started_at          timestamp(6) with local time zone,
    completed_at        timestamp(6) with local time zone,
    last_error          varchar2(4000 char),
    notes               varchar2(4000 char),
    created_at          timestamp(6) with local time zone default systimestamp not null,
    updated_at          timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_operation primary key (operation_id),
    constraint uq_form_submit_operation unique (app_renew_id, operation_key),
    constraint ck_fso_configured check (configured in ('Y','N')),
    constraint ck_fso_operation_key check (operation_key in (
        'payment','files','records','pdf','emails','billing','custom_code'
    )),
    constraint ck_fso_source_scope check (source_scope in (
        'form_config','triggered','not_configured','not_needed','legacy_inferred'
    )),
    constraint ck_fso_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'
    ))
);

comment on column form_submit_operation.source_scope is 'form_config = form-wide default; triggered = field-conditional; legacy_inferred = Phase-1 derivation.';
comment on column form_submit_operation.state is 'Canonical vocabulary; see form_submit_tool.html OP_STATE.';

-- =======================================================================
-- form_submit_operation_item : item-level detail (each file / each email)
-- =======================================================================
create table form_submit_operation_item (
    item_id             number(12)            not null,
    operation_id        number(12)            not null,
    app_renew_id        number(12)            not null,
    operation_key       varchar2(50 char)     not null,
    item_key            varchar2(200 char)    not null,
    item_label          varchar2(500 char),
    item_target         varchar2(1000 char),
    state               varchar2(30 char)     not null,
    source_scope        varchar2(30 char),
    trigger_field       varchar2(100 char),
    trigger_operator    varchar2(20 char),
    trigger_value       varchar2(400 char),
    temp_location       varchar2(1000 char),
    permanent_location  varchar2(1000 char),
    content_hash        varchar2(128 char),
    size_bytes          number(12),
    attempts            number(5)             default 0 not null,
    max_attempts        number(5)             default 0 not null,
    next_attempt_at     timestamp(6) with local time zone,
    last_error          varchar2(4000 char),
    created_at          timestamp(6) with local time zone default systimestamp not null,
    updated_at          timestamp(6) with local time zone default systimestamp not null,
    constraint pk_form_submit_operation_item primary key (item_id),
    constraint uq_form_submit_operation_item unique (operation_id, item_key),
    constraint ck_fsoi_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'
    ))
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
create index ix_fsr_type_state_updated
    on form_submit_record (app_renew_type, durable_state, updated_at);

create index ix_fsr_retry
    on form_submit_record (durable_state, retry_after);

create index ix_fso_app_state
    on form_submit_operation (app_renew_id, state);

create index ix_fso_recovery
    on form_submit_operation (state, next_attempt_at, operation_key);

create index ix_fsoi_app_operation_state
    on form_submit_operation_item (app_renew_id, operation_key, state);

create index ix_fse_app_created
    on form_submit_event (app_renew_id, created_at);

-- =======================================================================
-- Foreign keys (no cascade: ledger is audit/recovery data)
-- =======================================================================
alter table form_submit_operation
    add constraint fk_fso_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

alter table form_submit_operation_item
    add constraint fk_fsoi_operation foreign key (operation_id)
    references form_submit_operation (operation_id);

alter table form_submit_operation_item
    add constraint fk_fsoi_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

alter table form_submit_event
    add constraint fk_fse_record foreign key (app_renew_id)
    references form_submit_record (app_renew_id);

-- Optional when app_renew FK is allowed in the deployment:
-- alter table form_submit_record
--     add constraint fk_fsr_app_renew foreign key (app_renew_id)
--     references app_renew (app_renew_id);

-- =======================================================================
-- v_form_submit_admin : flattened view for apprenewadmin.asp and dashboards
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
    max(case when o.operation_key = 'payment' then o.state end)       payment_state,
    max(case when o.operation_key = 'files' then o.state end)          files_state,
    max(case when o.operation_key = 'files' then o.done_count end)    files_done_count,
    max(case when o.operation_key = 'files' then o.expected_count end) files_expected_count,
    max(case when o.operation_key = 'pdf' then o.state end)           pdf_state,
    max(case when o.operation_key = 'emails' then o.state end)        emails_state,
    max(case when o.operation_key = 'emails' then o.done_count end)   emails_done_count,
    max(case when o.operation_key = 'emails' then o.expected_count end) emails_expected_count,
    max(case when o.operation_key = 'records' then o.state end)       records_state,
    (select count(*) from form_submit_event e
       where e.app_renew_id = r.app_renew_id
         and e.event_type = 'awaiting_updates_released')               released_count
from form_submit_record r
left join form_submit_operation o on o.app_renew_id = r.app_renew_id
group by
    r.app_renew_id,
    r.app_renew_type,
    r.membership_id,
    r.durable_state,
    r.legacy_app_renew_status,
    r.ready_to_process,
    r.submitted_at,
    r.completed_at,
    r.updated_at;

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
